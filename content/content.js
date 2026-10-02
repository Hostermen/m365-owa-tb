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

  // --- Inject page-hook.js into the page world to intercept MSAL's fetch/XHR ---
  // The page hook patches fetch() and XMLHttpRequest in the PAGE world (where MSAL runs)
  // to capture token endpoint requests and responses. This is the only reliable way
  // to see the exact parameters MSAL sends and to capture refresh tokens from the
  // token endpoint response, since the webRequest API may not observe page-originated
  // requests in Thunderbird.
  try {
    var script = document.createElement("script");
    script.src = browser.runtime.getURL("content/page-hook.js");
    script.async = false;
    (document.head || document.documentElement).appendChild(script);
    script.onload = function() { script.remove(); };
    console.log("[M365OWA] content: injected page-hook.js into page world");
  } catch (e) {
    console.warn("[M365OWA] content: page-hook injection failed:", e.message || e);
  }

  // Listen for token response messages from the page hook (page world)
  window.addEventListener("message", function(event) {
    if (event.source !== window) return;
    var data = event.data;
    if (!data || data.type !== "M365OWA_TOKEN_RESPONSE") return;
    console.log("[M365OWA] content: received page-hook token response" +
      " (refresh_token=" + (data.refresh_token ? "yes(len=" + data.refresh_token.length + ")" : "no") +
      " client_id=" + (data.client_id || "?").slice(0, 12) +
      " url=" + (data.url || "?") + ")");
    try {
      browser.runtime.sendMessage({
        type: "m365-owa-refresh-token-harvest",
        refresh_tokens: [{
          refreshToken: data.refresh_token,
          clientId: data.client_id || OWA_CLIENT_ID,
          scope: data.scope || null,
          url: data.url || null,
          environment: null,
        }],
        access_token: data.access_token || null,
        source: "page-hook",
      }).catch(function() {});
    } catch (e) {
      console.warn("[M365OWA] content: page-hook relay failed:", e.message || e);
    }
  });

  // OWA's well-known public client_id (Microsoft Office / outlook).
  var OWA_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";

  // Fingerprint of the last relayed token set (avoid spamming on every poll).
  var lastFingerprint = null;
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
    var refreshTokens = [];  // array of {refreshToken, clientId, scope, url, environment}
    var accessToken = null;
    var globalScope = null, globalRealm = null, globalEnv = null;
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
        var keyParts = lk.split("|");
        // extract tenant/realm from this entry's own key (homeAccountId = {oid}.{tenantId})
        var entryRealm = null;
        if (keyParts.length >= 2) {
          var homeAcct = keyParts[1];
          var dotIdx = homeAcct.indexOf(".");
          if (dotIdx !== -1) entryRealm = homeAcct.slice(dotIdx + 1);
          if (entryRealm && !globalRealm) globalRealm = entryRealm;
        }
        // extract environment (authority hostname) from keyParts[2]
        // e.g., "login.windows.net" or "login.microsoftonline.com"
        var entryEnv = null;
        if (keyParts.length >= 3 && keyParts[2]) {
          entryEnv = keyParts[2];
          if (!globalEnv) globalEnv = entryEnv;
        }
        // extract scope from the accesstoken key (keyParts[6] = target/scope)
        if (credType === "accesstoken" && keyParts.length >= 7) {
          var scopeFromKey = keyParts.slice(6).join("|");
          try { scopeFromKey = decodeURIComponent(scopeFromKey); } catch (e) {}
          if (scopeFromKey) scopeFromKey = scopeFromKey.replace(/\|+$/, "");
          if (scopeFromKey && (!globalScope || scopeFromKey.indexOf("outlook") !== -1)) {
            globalScope = scopeFromKey;
          }
        }
        try { val = store.getItem(key); } catch (e) { continue; }
        if (!val) continue;
        var obj;
        try { obj = JSON.parse(val); } catch (e) { continue; }
        if (!obj || typeof obj !== "object") continue;
        // the raw token is in the `secret` field (MSAL v5) or `data` field (older format)
        var tok = obj.secret || obj.data;
        if (!tok || typeof tok !== "string") continue;
        // also check obj.target for scope (MSAL v3 standard cache value format)
        if (obj.target && (!globalScope || obj.target.indexOf("outlook") !== -1)) {
          globalScope = obj.target;
        }
        if (credType === "refreshtoken") {
          // Azure AD v2 refresh tokens typically start with "0.A" or "1.A" patterns
          var validTokenFormat = tok.length > 100 && (tok.indexOf("1.A") === 0 || tok.indexOf("0.A") === 0);
          // extract the real clientId from the MSAL cache key (keyParts[4])
          // format: msal.3|{homeAccountId}|{environment}|{credentialType}|{clientId}|{realm}|{target}
          var entryClientId = (keyParts.length >= 5 && keyParts[4] && keyParts[4].length > 10) ? keyParts[4] : OWA_CLIENT_ID;
          // log EVERY refreshtoken entry found (always-on diagnostic)
          console.log("[M365OWA] content: REFRESHTOKEN entry" +
            " (len=" + tok.length +
            " field=" + (obj.secret ? "secret" : "data") +
            " valid=" + validTokenFormat +
            " start=" + tok.slice(0, 20) +
            " end=..." + tok.slice(-15) +
            " clientId=" + entryClientId.slice(0, 12) +
            " hasOWA=" + (lk.indexOf(OWA_CLIENT_ID) !== -1) +
            " env=" + (entryEnv || "?") +
            " key=" + lk.slice(0, 120) + ")");
          // collect ALL refresh tokens — the Azure AD token endpoint is the
          // only real validator, so we don't filter by format here. The
          // background tries each one at renewal time.
          var realmForEntry = entryRealm || globalRealm;
          var envForEntry = entryEnv || globalEnv || "login.microsoftonline.com";
          // build the token endpoint URL using the ACTUAL authority hostname
          // from the MSAL cache key (login.windows.net or login.microsoftonline.com)
          var tokenUrl = realmForEntry
            ? "https://" + envForEntry + "/" + realmForEntry + "/oauth2/v2.0/token"
            : null;
          refreshTokens.push({
            refreshToken: tok,
            clientId: entryClientId,
            scope: globalScope,
            url: tokenUrl,
            environment: envForEntry,
          });
        } else if (credType === "accesstoken" && lk.indexOf(OWA_CLIENT_ID) !== -1) {
          // only keep JWT access tokens (start with "eyJ")
          if (tok.indexOf("eyJ") === 0 && !accessToken) {
            accessToken = tok;
          }
        }
      }
    }
    if (!refreshTokens.length && !accessToken) return null;
    return {
      refreshTokens: refreshTokens,
      accessToken: accessToken,
    };
  }

  // Relay captured tokens to the background script, but only when the set
  // actually changed since the last relay.
  function relay(found) {
    // fingerprint the token set for change detection
    var fp = found.refreshTokens.map(function(e) {
      return e.refreshToken.slice(0, 20) + ":" + e.refreshToken.length;
    }).join("|") + "|access=" + (found.accessToken ? found.accessToken.slice(0, 20) : "none");
    if (fp === lastFingerprint) return;
    lastFingerprint = fp;
    console.log("[M365OWA] content: relaying " + found.refreshTokens.length + " refresh token(s)" +
      " + " + (found.accessToken ? "access" : "no access") + " to background");
    try {
      browser.runtime.sendMessage({
        type: "m365-owa-refresh-token-harvest",
        refresh_tokens: found.refreshTokens,
        access_token: found.accessToken || null,
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
          // expired access token — drop it, keep the refresh tokens
          found.accessToken = null;
          // also forget our fingerprint so a fresh one relay later
          lastFingerprint = null;
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
