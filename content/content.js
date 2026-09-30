// Content script (isolated world): reads MSAL.js's token cache from the
// page's sessionStorage/localStorage and relays the access token + refresh
// token to the background script so auth survives Thunderbird restarts.
//
// Why a cache scan (instead of patching fetch/XHR): OWA's MSAL.js frequently
// reuses tokens already present in its browser cache (no token-endpoint POST
// at all), and when it does refresh it can do so from a worker/iframe that
// page-world fetch/XHR hooks cannot observe. MSAL ALWAYS writes every
// credential (access token, refresh token, id token) to its cache, so reading
// the cache is the only mechanism that reliably captures the refresh token
// in every code path. This also avoids patching the page's network APIs,
// which previously risked interfering with OWA's own scripts.
(function() {
  console.log("[M365OWA] content script loaded on:", location.href);

  // OWA's well-known public client_id (Microsoft Office / outlook).
  var OWA_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";

  // Last values relayed to the background (avoid spamming on every poll).
  var lastRefresh = null;
  var lastAccess = null;
  // Whether debug mode is enabled (checked once from extension storage).
  // Toggle from the Browser Toolbox console:
  //   browser.storage.local.set({ m365_owa_debug: true })
  //   browser.storage.local.remove("m365_owa_debug")
  var debugMode = false;
  var diagDone = false;
  try {
    browser.storage.local.get("m365_owa_debug").then(function(r) {
      debugMode = !!r.m365_owa_debug;
    }).catch(function() {});
    browser.storage.onChanged.addListener(function(changes, area) {
      if (area === "local" && changes.m365_owa_debug) {
        debugMode = !!changes.m365_owa_debug.newValue;
        if (debugMode) diagDone = false;
      }
    });
  } catch (e) {}

  // Return the page-world Storage object (read through wrappedJSObject so we
  // always read the PAGE's storage, where MSAL wrote the credentials, never an
  // isolated copy). Returns null when storage is unavailable or blocked.
  function getStore(which) {
    try {
      var win = window.wrappedJSObject || window;
      var s = (which === "local") ? win.localStorage : win.sessionStorage;
      if (!s) return null;
      // touch length to confirm cross-origin access is allowed
      var n = s.length;
      return (typeof n === "number") ? s : null;
    } catch (e) {
      return null;
    }
  }

  // Decode the JWT exp claim (ms epoch) from a token; returns null when not a
  // parseable JWT or when there is no exp claim.
  function jwtExpMs(tok) {
    if (!tok || typeof tok !== "string") return null;
    var parts = tok.split(".");
    if (parts.length < 2) return null;
    try {
      // base64url -> base64 -> JSON
      var b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      // pad to a multiple of 4
      while (b64.length % 4) b64 += "=";
      var json = atob(b64);
      var payload = JSON.parse(json);
      return payload && payload.exp ? payload.exp * 1000 : null;
    } catch (e) {
      return null;
    }
  }

  // Scan MSAL's cache for a refresh token and an outlook-scoped access token.
  // MSAL v3 (msal-browser 5.x) cache format:
  //   Key:   msal.3|{homeAccountId}|{environment}|{credentialType}|{clientId}...
  //          where credentialType is "accesstoken", "refreshtoken", or "idtoken"
  //          homeAccountId = {userObjectId}.{tenantId}
  //   Value: {id, nonce, data, lastUpdatedAt}
  //          where `data` is the raw token string
  function scanMsalCache() {
    var refreshToken = null, accessToken = null;
    var clientId = null, scope = null, realm = null;
    var stores = [getStore("session"), getStore("local")];
    for (var si = 0; si < stores.length; si++) {
      var store = stores[si];
      if (!store) continue;
      var n;
      try { n = store.length; } catch (e) { continue; }
      for (var i = 0; i < n; i++) {
        var key, val;
        try { key = store.key(i); } catch (e) { continue; }
        if (!key) continue;
        var lk = String(key);
        // only inspect keys that look like MSAL credential entries
        if (lk.indexOf("msal") === -1) continue;
        // determine credential type from the key
        var credType = null;
        if (lk.indexOf("|refreshtoken|") !== -1) credType = "refreshtoken";
        else if (lk.indexOf("|accesstoken|") !== -1) credType = "accesstoken";
        else if (lk.indexOf("|idtoken|") !== -1) credType = "idtoken";
        if (!credType) continue;
        // extract tenant/realm from homeAccountId in the key (part after the dot)
        // key format: msal.3|{userObjectId}.{tenantId}|{env}|{credType}|...
        // for accesstoken: ...|{credType}|{clientId}|{realm}|{target}
        var keyParts = lk.split("|");
        if (keyParts.length >= 2 && !realm) {
          var homeAcct = keyParts[1];
          var dotIdx = homeAcct.indexOf(".");
          if (dotIdx !== -1) realm = homeAcct.slice(dotIdx + 1);
        }
        // extract scope from the accesstoken key (keyParts[6] = target/scope)
        // but prefer obj.target from the value (more reliable, no delimiter leakage)
        if (credType === "accesstoken" && keyParts.length >= 7) {
          var scopeFromKey = keyParts.slice(6).join("|");
          try { scopeFromKey = decodeURIComponent(scopeFromKey); } catch (e) {}
          // strip trailing pipe (key delimiter leakage)
          if (scopeFromKey) scopeFromKey = scopeFromKey.replace(/\|+$/, "");
          if (scopeFromKey && (!scope || scopeFromKey.indexOf("outlook") !== -1)) {
            scope = scopeFromKey;
          }
        }
        try { val = store.getItem(key); } catch (e) { continue; }
        if (!val) continue;
        var obj;
        try { obj = JSON.parse(val); } catch (e) { continue; }
        if (!obj || typeof obj !== "object") continue;
        // the raw token is in the `data` field (MSAL v5) or `secret` field (MSAL v3/v4)
        var tok = obj.data || obj.secret;
        if (!tok || typeof tok !== "string") continue;
        // prefer obj.target from the value over key-based extraction (no delimiter issues)
        if (obj.target && (!scope || obj.target.indexOf("outlook") !== -1)) {
          scope = obj.target;
        }

        if (credType === "refreshtoken") {
          // prefer the refresh token for the OWA client_id
          if (!refreshToken || lk.indexOf(OWA_CLIENT_ID) !== -1) {
            refreshToken = tok;
            // extract client_id from the key (after |refreshtoken|)
            if (lk.indexOf(OWA_CLIENT_ID) !== -1) clientId = OWA_CLIENT_ID;
            // log token details for debugging (first/last 15 chars only)
            console.log("[M365OWA] content: refresh token found" +
              " (len=" + tok.length +
              " field=" + (obj.data ? "data" : "secret") +
              " start=" + tok.slice(0, 15) + "..." +
              " end=..." + tok.slice(-15) +
              " key=" + lk.slice(0, 80) + ")");
          }
        } else if (credType === "accesstoken" && lk.indexOf(OWA_CLIENT_ID) !== -1) {
          // only keep JWT access tokens (start with "eyJ")
          if (tok.indexOf("eyJ") === 0 && !accessToken) {
            accessToken = tok;
          }
        }
      }
    }
    if (!refreshToken && !accessToken) return null;
    // build the token endpoint URL from the tenant realm
    var url = null;
    if (realm) url = "https://login.microsoftonline.com/" + realm + "/oauth2/v2.0/token";
    return {
      refreshToken: refreshToken,
      accessToken: accessToken,
      clientId: clientId,
      scope: scope,
      url: url,
    };
  }

  // Relay captured tokens to the background script, but only when a value
  // actually changed since the last relay.
  function relay(found) {
    var changed = false;
    if (found.refreshToken && found.refreshToken !== lastRefresh) {
      lastRefresh = found.refreshToken;
      changed = true;
    }
    if (found.accessToken && found.accessToken !== lastAccess) {
      lastAccess = found.accessToken;
      changed = true;
    }
    if (!changed) return;
    console.log("[M365OWA] content: relaying MSAL cache tokens to background " +
      "(refresh=" + (found.refreshToken ? "yes" : "no") +
      ", access=" + (found.accessToken ? "yes" : "no") + ")");
    try {
      browser.runtime.sendMessage({
        type: "m365-owa-refresh-token-harvest",
        refresh_token: found.refreshToken || null,
        access_token: found.accessToken || null,
        client_id: found.clientId || null,
        scope: found.scope || null,
        url: found.url || null,
      }).catch(function() {});
    } catch (e) {
      console.warn("[M365OWA] content: relay failed:", e.message || e);
    }
  }

  // One-time diagnostic dump of MSAL cache entries (only when debug mode is on).
  // Toggle debug mode from the Browser Toolbox console:
  //   browser.storage.local.set({ m365_owa_debug: true })
  function diagDump() {
    if (!debugMode || diagDone) return;
    diagDone = true;
    var stores = [["sessionStorage", getStore("session")], ["localStorage", getStore("local")]];
    for (var si = 0; si < stores.length; si++) {
      var label = stores[si][0], store = stores[si][1];
      if (!store) { console.log("[M365OWA] DIAG " + label + ": NOT accessible"); continue; }
      var n;
      try { n = store.length; } catch (e) { continue; }
      console.log("[M365OWA] DIAG " + label + ": " + n + " keys total");
      for (var i = 0; i < n; i++) {
        var key;
        try { key = store.key(i); } catch (e) { continue; }
        if (!key) continue;
        var lk = String(key);
        if (lk.indexOf("msal") === -1) continue;
        var val = null;
        try { val = store.getItem(key); } catch (e) { continue; }
        var summary = "(no value)";
        if (val) {
          try {
            var obj = JSON.parse(val);
            var dType = typeof obj.data;
            var dLen = (typeof obj.data === "string") ? obj.data.length : "?";
            summary = "keys=[" + Object.keys(obj).join(",") + "] data:" + dType + " len=" + dLen;
          } catch (e2) {
            summary = "raw len=" + val.length;
          }
        }
        console.log("[M365OWA] DIAG " + label + " key[" + i + "]: " + lk.slice(0, 120) + " => " + summary);
      }
    }
  }

  // One scan + relay pass. Skips relaying an already-expired access token so we
  // never persist a stale token as "valid" (the refresh token still goes out
  // and drives the background renewal on its own).
  function tick() {
    try {
      diagDump();
      var found = scanMsalCache();
      if (!found) return;
      if (found.accessToken) {
        var exp = jwtExpMs(found.accessToken);
        if (exp != null && exp <= Date.now()) {
          // expired access token — drop it, keep the refresh token
          found.accessToken = null;
          // also forget our "last seen" so a fresh one relay later
          lastAccess = null;
        }
      }
      relay(found);
    } catch (e) {
      console.warn("[M365OWA] content: scan failed:", e.message || e);
    }
  }

  // Scan now, then poll. MSAL populates the cache shortly after the page
  // loads (or immediately when reusing a cached session), so frequent early
  // polling catches a fresh login while slower later polling keeps the
  // tokens current as MSAL rotates them.
  tick();
  var fast = setInterval(tick, 3000);
  setTimeout(function() {
    clearInterval(fast);
    setInterval(tick, 30000);
  }, 120000);

  // The interactive login returns to a URL whose hash carries `#code=...`;
  // MSAL exchanges that code and writes the cache shortly after, so re-scan
  // on a short schedule when we land on such a URL.
  if (location.hash && location.hash.indexOf("code=") !== -1) {
    setTimeout(tick, 1500);
    setTimeout(tick, 4000);
    setTimeout(tick, 10000);
  }
})();
