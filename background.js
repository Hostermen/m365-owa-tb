// Startup entry point: loads config and token, registers the token harvester, then bootstraps contacts (best-effort) and the calendar provider.
(async function() {
  // load runtime config overrides from storage
  await loadConfig();
  // restore the stored bearer token
  await Auth.loadStoredToken();
  // restore the stored refresh token (survives reboots, no SSO dependency)
  await Auth.loadStoredRefreshToken();
  // restore the PKCE refresh token (90-day rolling — primary auth method)
  await Auth.loadStoredPkceToken();
  // register the OWA Authorization-header harvester
  registerHarvester();
  // register the refresh-token harvester on login.microsoftonline.com
  registerRefreshTokenHarvester();
  // schedule the periodic token-renewal alarm
  await browser.alarms.create("m365-owa-renew", { periodInMinutes: 20 });
  // schedule a connectivity-probe alarm (every 30s) so the offline → online
  // transition is detected quickly even when the options page is closed.
  // navigator.onLine/window events are unreliable in the TB background.
  await browser.alarms.create("m365-owa-netprobe", { periodInMinutes: 0.5 });
  // probe connectivity once at startup so _netOnline is accurate before the
  // first maybeRenew() (otherwise a cold boot with no internet would try the
  // slow OWA frame/tab fallbacks before the probe marks the net offline).
  await _probeConnectivity(true);
  // log the startup auth state and configured host
  console.log("[M365OWA] startup. authenticated:", Auth.isAuthenticated(), "host:", CONFIG.OWA_HOST, "online:", _netOnline);

  // if already authenticated, initialise contacts sync (best-effort)
  if (Auth.isAuthenticated()) {
    try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
  }
  // initialise the calendar sync provider (self-bootstraps on first pull)
  CalendarSync.init();
  // opportunistically renew an expired/stale token left over from a previous session (cold-boot recovery)
  maybeRenew().catch((e) => console.warn("[M365OWA] startup renewal failed:", e.message || e));
  // verify the token works against OWA once at startup (after the renewal)
  _runProbe().catch(() => {});
})();

// OWA service endpoints whose Authorization headers carry fresh bearer tokens.
function owaServiceUrls() {
  // return one match pattern per first-party OWA host
  return [
    "https://outlook.office.com/owa/service.svc*",
    "https://outlook.office365.com/owa/service.svc*",
    "https://outlook.cloud.microsoft/owa/service.svc*",
  ];
}

// Observe OWA service requests and harvest bearer tokens from their Authorization headers.
function registerHarvester() {
  // listen for outgoing OWA service requests with headers visible
  browser.webRequest.onBeforeSendHeaders.addListener(
    // harvest the Authorization header from each request
    (details) => {
      // skip requests without headers
      if (!details.requestHeaders) return;
      // trace service calls while a renewal is being diagnosed
      if (_renewalActive) {
        // note that OWA issued a service call
        _renewalSawServiceCall = true;
        // log the called endpoint for diagnostics
        console.log("[M365OWA] renewal: service call observed", details.url);
      }
      // scan all request headers
      for (const h of details.requestHeaders) {
        // match the Authorization header case-insensitively
        if (h.name && h.name.toLowerCase() === "authorization" && h.value) {
          // remember whether we were authenticated before this harvest
          const wasAuth = Auth.isAuthenticated();
          // store the harvested token (fire-and-forget promise)
          Auth.harvestFromHeader(h.value).then((ok) => {
            // a fresh token was captured — cancel any pending renewal retry
            if (ok) _clearRenewRetry();
            // bootstrap sync on the login transition only
            if (ok && !wasAuth && Auth.isAuthenticated()) onFirstHarvest();
          }).catch((e) => console.warn("[M365OWA] harvest failed:", e.message || e));
          // stop after the first Authorization header
          break;
        }
      }
    },
    // restrict observation to OWA service endpoints
    { urls: owaServiceUrls() },
    // request header visibility (observe only, never block)
    ["requestHeaders"]
  );
  console.log("[M365OWA] token harvester registered");
}

// Store request body data (client_id, scope) keyed by requestId for use when the response arrives.
const _tokenRequestData = new Map();

// Observe OWA's token endpoint calls and harvest refresh tokens from the RESPONSE body.
// Uses webRequest.filterResponseData to intercept the response (refresh_token is in the
// response for ALL grant types: authorization_code, refresh_token, etc.).
function registerRefreshTokenHarvester() {
  // check whether filterResponseData is available (Thunderbird 57+/Firefox 57+)
  const useFilter = typeof browser.webRequest.filterResponseData === "function";
  console.log("[M365OWA] filterResponseData available:", useFilter);

  // listen for POSTs to the OAuth2 token endpoints on login.microsoftonline.com
  browser.webRequest.onBeforeRequest.addListener(
    (details) => {
      // log that we saw a token endpoint request (diagnostic)
      console.log("[M365OWA] webRequest: token endpoint hit:", details.method, details.url);
      // skip our own refresh requests
      if (Auth._selfRefreshActive) return;
      // capture client_id and scope from the request body for later use with the response
      let cid = null, scope = null;
      if (details.requestBody && details.requestBody.formData) {
        const fd = details.requestBody.formData;
        cid = (fd.client_id && fd.client_id[0]) || null;
        scope = (fd.scope && fd.scope[0]) || null;
        // log ALL form fields for MSAL's own token requests (diagnostic)
        console.log("[M365OWA] webRequest: MSAL token request params:" +
          " grant_type=" + (fd.grant_type && fd.grant_type[0] || "?") +
          " client_id=" + (cid || "?").slice(0, 12) +
          " scope=" + (scope || "none").slice(0, 100) +
          " resource=" + (fd.resource && fd.resource[0] || "none") +
          " client_info=" + (fd.client_info && fd.client_info[0] || "none") +
          " refresh_token_len=" + (fd.refresh_token && fd.refresh_token[0] && fd.refresh_token[0].length || 0) +
          " code_len=" + (fd.code && fd.code[0] && fd.code[0].length || 0));
      }
      if (useFilter) {
        // store request data keyed by requestId so the response handler can use it
        _tokenRequestData.set(details.requestId, { cid, scope, url: details.url });
        // create a response filter to intercept the token endpoint response body
        const filter = browser.webRequest.filterResponseData(details.requestId);
        const decoder = new TextDecoder("utf-8");
        let responseText = "";
        // pass through each chunk unchanged and accumulate the response text
        filter.ondata = (event) => {
          filter.write(event.data);
          responseText += decoder.decode(event.data, { stream: true });
        };
        // when the response is complete, parse it and harvest the refresh token
        filter.onstop = () => {
          filter.close();
          const reqData = _tokenRequestData.get(details.requestId) || {};
          _tokenRequestData.delete(details.requestId);
          // log the token response for diagnostics
          console.log("[M365OWA] webRequest: token response (len=" + responseText.length + "): " + responseText.slice(0, 400));
          try {
            const data = JSON.parse(responseText);
            if (data.refresh_token) {
              console.log("[M365OWA] webRequest: refresh_token in response! len=" + data.refresh_token.length +
                " start=" + data.refresh_token.slice(0, 20) + " token_type=" + (data.token_type || "?") +
                " scope=" + (data.scope || "none").slice(0, 100));
              const wasAuth = Auth.isAuthenticated();
              Auth.harvestFromTokenResponse(
                data.refresh_token,
                data.access_token,
                reqData.cid || null,
                reqData.url || details.url,
                reqData.scope || data.scope || null,
                "webRequest"
              ).then((ok) => {
                if (ok) {
                  _clearRenewRetry();
                  if (!wasAuth && Auth.isAuthenticated()) onFirstHarvest();
                }
              }).catch((e) => console.warn("[M365OWA] refresh token harvest (response) failed:", e.message || e));
            } else {
              console.log("[M365OWA] webRequest: no refresh_token in response. Keys:", Object.keys(data).join(","));
            }
          } catch (e) {
            console.log("[M365OWA] webRequest: response parse failed:", e.message || e);
          }
        };
        filter.onerror = () => {
          try { filter.disconnect(); } catch {}
          _tokenRequestData.delete(details.requestId);
        };
      } else {
        // fallback: harvest from request body only (refresh_token grant)
        if (details.requestBody && details.requestBody.formData) {
          const fd = details.requestBody.formData;
          if (fd.grant_type && fd.grant_type[0] === "refresh_token" && fd.refresh_token && fd.refresh_token[0]) {
            Auth.harvestRefreshToken(fd, details.url)
              .catch((e) => console.warn("[M365OWA] refresh token harvest (request) failed:", e.message || e));
          }
        }
      }
    },
    // match the v2.0 and v1.0 token endpoints on login.microsoftonline.com AND login.windows.net
    { urls: [
      "https://login.microsoftonline.com/*/oauth2/v2.0/token",
      "https://login.microsoftonline.com/*/oauth2/token",
      "https://login.windows.net/*/oauth2/v2.0/token",
      "https://login.windows.net/*/oauth2/token",
    ] },
    // request body visibility + blocking (required by filterResponseData)
    useFilter ? ["requestBody", "blocking"] : ["requestBody"]
  );
  console.log("[M365OWA] refresh token harvester registered (mode: " + (useFilter ? "response-filter" : "request-only") + ")");
}

// --- PKCE OAuth2 login flow (90-day rolling refresh tokens, no Azure app registration) ---
// Pending PKCE login state: set when the login tab is opened, cleared when the
// redirect is intercepted or the login times out.
// Shape: { verifier, state, tabId, resolve, reject, timer }
let _pkcePending = null;

// Intercept the OAuth2 redirect to http://localhost and extract the auth code.
// Registered once at startup; active whenever a PKCE login is in progress.
browser.webRequest.onBeforeRequest.addListener(
  (details) => {
    // no pending login → ignore
    if (!_pkcePending) return;
    // parse the redirect URL
    let url;
    try { url = new URL(details.url); } catch { return; }
    // verify the state parameter matches our pending login (CSRF protection)
    const state = url.searchParams.get("state");
    if (state !== _pkcePending.state) return;
    // check for an error response (e.g. access_denied, invalid_scope)
    const error = url.searchParams.get("error");
    if (error) {
      const errorDesc = url.searchParams.get("error_description") || "";
      console.log("[M365OWA] PKCE: login error from Microsoft: " + error + " — " + errorDesc);
      _pkcePending.reject(new Error("OAuth error: " + error + " — " + errorDesc));
      clearTimeout(_pkcePending.timer);
      _pkcePending = null;
      return { cancel: true };
    }
    // extract the authorization code
    const code = url.searchParams.get("code");
    if (!code) {
      console.log("[M365OWA] PKCE: redirect has no code parameter");
      return { cancel: true };
    }
    console.log("[M365OWA] PKCE: auth code captured from redirect (len=" + code.length + ")");
    // close the login tab (no longer needed)
    if (_pkcePending.tabId != null) {
      try { browser.tabs.remove(_pkcePending.tabId); } catch {}
    }
    // resolve the pending login promise with the auth code
    _pkcePending.resolve(code);
    clearTimeout(_pkcePending.timer);
    _pkcePending = null;
    // cancel the request (prevent localhost from actually loading)
    return { cancel: true };
  },
  { urls: ["http://localhost/*"] },
  ["blocking"]
);
console.log("[M365OWA] PKCE redirect interceptor registered");

// Strip the Origin header from our own PKCE token requests so Microsoft's
// token endpoint treats them as native/public-client requests (not cross-origin
// SPA requests). Without this, AAD rejects the authorization_code redemption with
// AADSTS9002326 ("Cross-origin token redemption is permitted only for the
// 'Single-Page Application' client-type") because Thunderbird attaches
// Origin: moz-extension://<uuid> to background-page fetch() calls.
// OWA's own MSAL requests originate from content tabs (Origin: https://outlook.*)
// and are left untouched by the moz-extension:// filter below.
browser.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    let changed = false;
    const kept = [];
    for (const h of (details.requestHeaders || [])) {
      const name = (h.name || "").toLowerCase();
      // drop the Origin header only when our extension issued the request
      if (name === "origin" && /^(moz|chrome)-extension:/.test(h.value || "")) {
        changed = true;
        continue;
      }
      kept.push(h);
    }
    return changed ? { requestHeaders: kept } : {};
  },
  { urls: [
    "https://login.microsoftonline.com/*/oauth2/v2.0/token",
    "https://login.microsoftonline.com/*/oauth2/token",
    "https://login.windows.net/*/oauth2/v2.0/token",
    "https://login.windows.net/*/oauth2/token",
  ] },
  ["blocking", "requestHeaders"]
);
console.log("[M365OWA] PKCE Origin stripper registered");

// Start the PKCE OAuth2 login flow: open a Microsoft login tab, wait for the
// redirect, exchange the code for tokens, and store them. Returns { ok, reason }.
// The user logs in once; the resulting refresh token lasts 90 days (rolling).
async function pkceLogin() {
  // refuse to start if a login is already in progress
  if (_pkcePending) return { ok: false, reason: "PKCE login already in progress" };
  try {
    // build the authorization URL + PKCE parameters
    const { url, verifier, state } = await TbOAuth.buildAuthUrl();
    console.log("[M365OWA] PKCE: starting login flow (client_id=" + TbOAuth.CLIENT_ID.slice(0, 8) + "..." + ")");
    // open the Microsoft login tab
    const tab = await browser.tabs.create({ url, active: true });
    // create a promise that resolves when the redirect is intercepted
    const codePromise = new Promise((resolve, reject) => {
      // 5-minute timeout — user may walk away or close the tab
      const timer = setTimeout(() => {
        if (_pkcePending && _pkcePending.state === state) {
          console.log("[M365OWA] PKCE: login timed out (5 minutes)");
          _pkcePending.reject(new Error("Login timed out (5 minutes)"));
          _pkcePending = null;
        }
      }, 5 * 60 * 1000);
      _pkcePending = { verifier, state, tabId: tab.id, resolve, reject, timer };
    });
    // wait for the auth code (from the redirect interceptor)
    const code = await codePromise;
    // exchange the code for access + refresh tokens
    const data = await TbOAuth.exchangeCode(code, verifier);
    // store the tokens (access token + 90-day refresh token)
    await Auth.setPkceTokens(data);
    console.log("[M365OWA] PKCE: login flow complete — 90-day refresh token stored");
    // bootstrap contacts + calendar sync
    await onFirstHarvest();
    return { ok: true };
  } catch (e) {
    // clean up pending state on any error
    if (_pkcePending) {
      clearTimeout(_pkcePending.timer);
      _pkcePending = null;
    }
    console.warn("[M365OWA] PKCE login failed:", e.message || e);
    return { ok: false, reason: (e && e.message) || String(e) };
  }
}

// Bootstrap contacts and calendar once, right after the first token harvest of a session.
async function onFirstHarvest() {
  console.log("[M365OWA] first token harvested, bootstrapping sync");
  // initialise contacts sync (best-effort)
  try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
  // run calendar bootstrap so the provider calendar is created now that we have a token
  try { await CalendarSync._bootstrap(); } catch (e) { console.error("[M365OWA] calendar bootstrap failed", e); }
  // trigger a calendar sync
  await messenger.calendar.calendars.synchronize().catch(() => {});
}

// Forget the tracked OWA tab as soon as the user closes it.
browser.tabs.onRemoved.addListener((tabId) => {
  // clear the tracked id when our OWA tab was closed
  if (Auth._owaTabId === tabId) Auth._owaTabId = null;
});

// OWA page URLs loaded into the hidden renewal frame (host-wide: OWA redirects across /owa/, /mail/, /calendar/).
function owaPageUrls() {
  // return one match pattern per first-party OWA host
  return [
    "https://outlook.office.com/*",
    "https://outlook.office365.com/*",
    "https://outlook.cloud.microsoft/*",
  ];
}

// True while a hidden renewal attempt is running.
let _renewalActive = false;
// Set when an OWA frame response passes through during renewal.
let _renewalSawFrameResponse = false;
// Set when an OWA service call is observed during renewal.
let _renewalSawServiceCall = false;
// Handle of the pending renewal-retry setTimeout (null when none scheduled).
let _renewRetryTimer = null;
// Index into the renewal-retry backoff sequence for the current retry run.
let _renewRetryIndex = 0;
// True while any renewal attempt (PKCE / OWA refresh / hidden iframe / tab) is
// in progress. Exposed via status() so the options badge can show a pending
// state instead of "Not connected" while a background refresh is running.
let _renewInProgress = false;

// --- Network connectivity probe ---
// navigator.onLine is unreliable in the Thunderbird background context (it
// reports "online" whenever any network interface is up, even with no
// internet). We probe the Microsoft login endpoint directly: any HTTP
// response means we can reach Microsoft; a network error means offline.
// The result drives the options badge ("Offline ✗" in red) and triggers an
// immediate renewal on the offline → online transition.
let _netOnline = true;
let _netProbeTs = 0;
const NET_PROBE_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const NET_PROBE_TIMEOUT_MS = 5000;
const NET_PROBE_TTL_MS = 10000;

// Last OWA auth-probe result (a GetCalendarView call verifying the token
// works). Stored so status() can report it without re-probing on every 5s
// poll. Refreshed by _runProbe() on startup, network reconnection, and
// explicit diagnostics (Test connection).
let _owaProbeResult = null;

async function _runProbe(force) {
  if (!Auth.isAuthenticated()) { _owaProbeResult = null; return _owaProbeResult; }
  try {
    await OWA.probe(force);
    _owaProbeResult = "ok";
  } catch (e) {
    _owaProbeResult = "probe failed: " + (e.message || e);
  }
  return _owaProbeResult;
}

async function _probeConnectivity(force) {
  // return cached result when fresh (and not forced)
  if (!force && Date.now() - _netProbeTs < NET_PROBE_TTL_MS) return _netOnline;
  const prev = _netOnline;
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), NET_PROBE_TIMEOUT_MS);
    // GET + redirect:manual → a 302 resolves as an opaque-redirect response
    // (no body downloaded); a network failure rejects. Either way we learn
    // whether Microsoft is reachable.
    await fetch(NET_PROBE_URL, {
      method: "GET",
      cache: "no-store",
      signal: ctrl.signal,
      credentials: "omit",
      redirect: "manual",
    });
    clearTimeout(t);
    _netOnline = true;
  } catch {
    _netOnline = false;
  }
  _netProbeTs = Date.now();
  // offline → online transition: trigger an immediate renewal so recovery
  // is instant (don't wait for the next backoff retry or the 20-min alarm).
  if (!prev && _netOnline) {
    console.log("[M365OWA] connectivity probe: offline → online, triggering renewal");
    maybeRenew().catch((e) => console.warn("[M365OWA] post-online renewal failed:", e.message || e));
    // verify the token still works against OWA after recovering connectivity
    _runProbe().catch(() => {});
  }
  // push connectivity changes to any open options page so the badge updates
  // in real time instead of waiting for the 5s poll.
  if (prev !== _netOnline) _broadcastConnectivity(_netOnline);
  return _netOnline;
}

// Broadcast connectivity state to any open options page (push-based badge
// updates). runtime.sendMessage is not delivered to the sender's own
// context, so the background's onMessage listener won't loop on this.
// Rejects silently when no options page is listening (page closed).
function _broadcastConnectivity(online) {
  try {
    browser.runtime.sendMessage({ type: "m365-owa-connectivity", online: !!online }).catch(() => {});
  } catch { /* no receiving end */ }
}

// Strip framing protections from OWA responses while a hidden renewal runs (registered temporarily).
function stripFrameHeaders(details) {
  // note that an OWA frame response was seen
  _renewalSawFrameResponse = true;
  // track whether any framing header was actually removed
  let changed = false;
  // collect the headers to keep
  const kept = [];
  // scan all response headers
  for (const h of details.responseHeaders || []) {
    // normalise the header name
    const name = (h.name || "").toLowerCase();
    // drop clickjacking headers so the hidden frame may load OWA
    if (name === "x-frame-options" || name === "frame-options") { changed = true; continue; }
    // rewrite CSP without the frame-ancestors directive
    if (name === "content-security-policy") {
      // split, drop frame-ancestors, rejoin
      const v = String(h.value || "").split(";").map((s) => s.trim()).filter((s) => s && !/^frame-ancestors/i.test(s)).join("; ");
      // note when the directive was present
      if (v !== String(h.value || "")) changed = true;
      // keep the rewritten policy when non-empty
      if (v) kept.push({ name: h.name, value: v });
      // skip the original header
      continue;
    }
    // keep everything else untouched
    kept.push({ name: h.name, value: h.value });
  }
  // log whether framing headers were stripped for diagnostics
  console.log("[M365OWA] renewal: framing " + (changed ? "stripped for " : "absent on ") + details.url);
  // return the filtered headers
  return { responseHeaders: kept };
}

// Wait until the harvester stores a token (token age drops); returns true on harvest, false on timeout.
async function waitForHarvest(baselineAgeMs, timeoutMs) {
  // record the start time
  const start = Date.now();
  // poll until the timeout expires
  while (Date.now() - start < timeoutMs) {
    // a fresh harvest resets the token age below the baseline
    if (Auth.tokenAgeMs() < baselineAgeMs) return true;
    // wait 2 seconds between polls
    await new Promise((r) => setTimeout(r, 2000));
  }
  // report timeout
  return false;
}

// Renew the token in the background: load OWA in the hidden frame so its session mints a fresh token.
async function renewTokenHidden() {
  // never overlap two renewals
  if (_renewalActive) return { ok: false, reason: "renewal already running" };
  // never act without a previously captured token or refresh token (i.e. user never connected)
  if (!Auth._token && !Auth.hasRefreshToken()) return { ok: false, reason: "no token stored" };
  // mark the renewal as running
  _renewalActive = true;
  // reset the diagnostic flags
  _renewalSawFrameResponse = false;
  // reset the service-call flag
  _renewalSawServiceCall = false;
  try {
    // record the current token age as the harvest baseline
    const baseline = Auth.tokenAgeMs();
    // temporarily allow framing OWA inside the background page
    browser.webRequest.onHeadersReceived.addListener(stripFrameHeaders, { urls: owaPageUrls(), types: ["sub_frame"] }, ["blocking", "responseHeaders"]);
    // locate the hidden renewal frame
    const frame = document.getElementById("owaRenewal");
    // load OWA with the stored session cookies
    frame.src = Auth.owaLoginUrl();
    console.log("[M365OWA] hidden renewal started");
    // wait up to 90 seconds for the harvester to capture a fresh token
    const ok = await waitForHarvest(baseline, 90000);
    // diagnose the failure stage when nothing was harvested
    let reason = "renewed";
    // no OWA response means the frame never loaded (framing still blocked or no session)
    if (!ok && !_renewalSawFrameResponse) reason = "hidden frame got no OWA response (framing blocked?)";
    // response without service calls means OWA rendered no session (login page?)
    else if (!ok && !_renewalSawServiceCall) reason = "OWA loaded but issued no service calls (not logged in?)";
    // service calls without a token mean unusable Authorization headers
    else if (!ok) reason = "service calls seen but no usable token captured";
    // log the outcome
    console.log("[M365OWA] hidden renewal " + (ok ? "succeeded" : "timed out: " + reason));
    // report the outcome
    return { ok, reason };
  } finally {
    // restore OWA framing protections immediately
    try { browser.webRequest.onHeadersReceived.removeListener(stripFrameHeaders); } catch {}
    // unload OWA from the hidden frame to free resources
    const frame = document.getElementById("owaRenewal");
    // navigate the frame away
    if (frame) frame.src = "about:blank";
    // mark the renewal as finished
    _renewalActive = false;
  }
}

// Renew the token by opening OWA in a real (background) tab; more reliable than the hidden iframe on cold boot.
async function renewTokenViaTab() {
  // never overlap two renewals
  if (_renewalActive) return { ok: false, reason: "renewal already running" };
  // never act without a previously captured token or refresh token (i.e. user never connected)
  if (!Auth._token && !Auth.hasRefreshToken()) return { ok: false, reason: "no token stored" };
  // mark the renewal as running
  _renewalActive = true;
  // remember the tab we open so it can be closed in finally
  let tabId = null;
  try {
    // record the current token age as the harvest baseline
    const baseline = Auth.tokenAgeMs();
    // open OWA in a background tab (active:false) using the stored session cookies
    const tab = await browser.tabs.create({ url: Auth.owaLoginUrl(), active: false });
    // remember the tab id for the onRemoved listener and cleanup
    tabId = tab.id;
    Auth._owaTabId = tabId;
    console.log("[M365OWA] tab renewal started, tab", tabId);
    // wait up to 90 seconds for the harvester to capture a fresh token
    const ok = await waitForHarvest(baseline, 90000);
    // diagnose the outcome
    const reason = ok ? "renewed" : "no token harvested from OWA tab (session cookies expired?)";
    // log the outcome
    console.log("[M365OWA] tab renewal " + (ok ? "succeeded" : "timed out: " + reason));
    // report the outcome
    return { ok, reason };
  } finally {
    // close the renewal tab if it is still open
    if (tabId != null) {
      try { await browser.tabs.remove(tabId); } catch {}
      // clear the tracked tab id when it was ours
      if (Auth._owaTabId === tabId) Auth._owaTabId = null;
    }
    // mark the renewal as finished
    _renewalActive = false;
  }
}

// True while a retry-with-backoff loop is scheduled (prevents overlapping retry timers).
// (_renewRetryTimer / _renewRetryIndex are declared with the other module-level state above.)

// Renew the token when it is missing, flagged expired, or older than 30 minutes.
// Tries the PKCE refresh token first (90-day rolling, most reliable), then the
// OWA-harvested refresh tokens (24h SPA), then the non-intrusive hidden iframe,
// then a background OWA tab, and on failure schedules a backoff retry.
async function maybeRenew() {
  // skip when no token was ever captured AND no refresh tokens of any kind stored
  if (!Auth._token && !Auth.hasRefreshToken() && !Auth.hasPkceRefreshToken()) return;
  // skip fresh, valid tokens — nothing to do
  if (Auth._token && !Auth._expired && Auth.tokenAgeMs() <= 30 * 60 * 1000) return;

  // remember whether we were authenticated before this renewal attempt
  // (cold boot: refresh token present but no live access token → bootstrap after)
  const wasAuth = Auth.isAuthenticated();
  // mark a renewal as in progress (for the options badge pending state)
  _renewInProgress = true;
  try {

  // === Priority 1: PKCE refresh token (90-day rolling, public client — never re-login) ===
  if (Auth.hasPkceRefreshToken()) {
    console.log("[M365OWA] maybeRenew: trying PKCE refresh token (90-day rolling)");
    const r = await Auth.refreshViaPkce();
    if (r.ok) {
      _clearRenewRetry();
      if (!wasAuth && Auth.isAuthenticated()) {
        onFirstHarvest().catch((e) => console.warn("[M365OWA] post-PKCE-renew bootstrap failed:", e.message || e));
      }
      return;
    }
    console.log("[M365OWA] PKCE refresh failed (" + r.reason + "), falling back to OWA refresh tokens");
  }

  // === Priority 2: OWA-harvested refresh tokens (24h SPA, less reliable) ===
  if (Auth.hasRefreshToken()) {
    console.log("[M365OWA] maybeRenew: trying OWA refresh tokens (count=" + Auth._refreshTokens.length +
      ", first url=" + (Auth._refreshTokens[0] ? Auth._refreshTokens[0].url || "none" : "none") + ")");
    const r = await Auth.refreshViaRefreshToken();
    if (r.ok) {
      _clearRenewRetry();
      // bootstrap contacts/calendar when this renewal flipped us from
      // unauthenticated → authenticated (typical cold-boot recovery)
      if (!wasAuth && Auth.isAuthenticated()) {
        onFirstHarvest().catch((e) => console.warn("[M365OWA] post-renew bootstrap failed:", e.message || e));
      }
      // SPA refresh tokens die 24h after the last interactive auth (OWA is a SPA).
      // Proactively trigger a cookie-based renewal (hidden iframe) when the
      // interactive harvest is stale, so OWA's MSAL mints a fresh SPA refresh
      // token with a new 24h window. Without this, the refresh token dies after
      // 24h and the addon falls back to cookie-only renewal (which fails if the
      // session cookie also expired, forcing the user to re-login).
      const SPA_REFRESH_INTERVAL = 12 * 60 * 60 * 1000; // 12h (half the 24h window)
      if (Auth._lastInteractiveHarvestTs &&
          Date.now() - Auth._lastInteractiveHarvestTs > SPA_REFRESH_INTERVAL) {
        const h = Math.round((Date.now() - Auth._lastInteractiveHarvestTs) / 3600000);
        console.log("[M365OWA] SPA token window stale (" + h + "h since last interactive harvest), triggering background cookie renewal");
        // defer 5s so maybeRenew() returns first and _renewalActive is free
        setTimeout(() => {
          renewTokenHidden().then((rr) => {
            if (rr.ok) console.log("[M365OWA] SPA proactive renewal succeeded (fresh refresh token harvested)");
            else console.log("[M365OWA] SPA proactive renewal failed: " + rr.reason);
          }).catch((e) => console.warn("[M365OWA] SPA proactive renewal error:", e.message || e));
        }, 5000);
      }
      return;
    }
    console.log("[M365OWA] OWA refresh token renewal failed (" + r.reason + "), falling back to hidden iframe");
  } else {
    console.log("[M365OWA] maybeRenew: no OWA refresh token stored, trying hidden iframe");
  }

  // When the network is unreachable, the hidden-iframe and tab fallbacks
  // cannot succeed (they need to load OWA over the network) and would just
  // burn 90s+90s before timing out. Skip them and go straight to a backoff
  // retry — the connectivity probe (_probeConnectivity) triggers an immediate
  // maybeRenew() as soon as Microsoft is reachable again.
  if (!_netOnline) {
    console.log("[M365OWA] network offline — skipping OWA frame/tab renewal, scheduling retry");
    _scheduleRenewRetry();
    return;
  }

  // fall back to the non-intrusive hidden iframe (works when the network is warm)
  let r = await renewTokenHidden();
  if (r.ok) { _clearRenewRetry(); return; }

  // fall back to a real OWA tab when the hidden iframe fails (e.g. cold boot)
  console.log("[M365OWA] hidden renewal failed (" + r.reason + "), falling back to OWA tab");
  r = await renewTokenViaTab();
  if (r.ok) { _clearRenewRetry(); return; }

  // all renewal methods failed — schedule a backoff retry so we keep trying automatically
  _scheduleRenewRetry();
  } finally {
    _renewInProgress = false;
  }
}

// Schedule a renewal retry with exponential backoff (1, 2, 4, 5 min), capped at 5 minutes.
function _scheduleRenewRetry() {
  // cancel any pending retry first
  _clearRenewRetry();
  // backoff sequence in ms: 1min, 2min, 4min, then stay at 5min
  const backoff = [60 * 1000, 2 * 60 * 1000, 4 * 60 * 1000, 5 * 60 * 1000];
  // pick the delay for this retry from the current index (capped at the last entry)
  const delay = backoff[Math.min(_renewRetryIndex || 0, backoff.length - 1)];
  // advance the retry index for the next attempt (capped at the last entry)
  _renewRetryIndex = Math.min((_renewRetryIndex || 0) + 1, backoff.length - 1);
  console.log("[M365OWA] renewal failed; retrying in " + (delay / 1000) + "s (attempt " + (_renewRetryIndex + 1) + ")");
  // schedule the retry
  _renewRetryTimer = setTimeout(() => {
    _renewRetryTimer = null;
    maybeRenew().catch((e) => console.warn("[M365OWA] retry renewal failed:", e.message || e));
  }, delay);
}

// Clear any pending renewal retry and reset the retry index.
function _clearRenewRetry() {
  // clear the timer when one is pending
  if (_renewRetryTimer) { clearTimeout(_renewRetryTimer); _renewRetryTimer = null; }
  // reset the backoff index
  _renewRetryIndex = 0;
}

// Periodic renewal alarm: re-check and renew the token every 20 minutes (also catches expiry between retries).
browser.alarms.onAlarm.addListener(async (alarm) => {
  // only handle our own alarms
  if (!alarm) return;
  if (alarm.name === "m365-owa-renew") {
    // run the renewal check, logging failures
    await maybeRenew().catch((e) => console.warn("[M365OWA] background renewal failed:", e.message || e));
  } else if (alarm.name === "m365-owa-netprobe") {
    // connectivity probe — detects offline → online transitions (and
    // triggers an immediate renewal) even when the options page is closed.
    await _probeConnectivity(true).catch(() => {});
  }
});

// Network connectivity listeners: trigger an immediate renewal as soon as the
// browser regains connectivity. On a cold boot with no network the token refresh
// fails and falls into the backoff retry loop; without this the addon would
// wait up to 5 minutes (next backoff tick) before noticing the network is back.
// The "online" event fires on the background page window when the OS reports a
// working network interface. (navigator.onLine is also used in maybeRenew() to
// skip the slow OWA frame/tab fallbacks while offline.)
window.addEventListener("online", () => {
  console.log("[M365OWA] network online — triggering renewal");
  // force a probe to confirm reachability, then push the result to the UI
  _probeConnectivity(true).then(() => _broadcastConnectivity(_netOnline)).catch(() => {});
  maybeRenew().catch((e) => console.warn("[M365OWA] online renewal failed:", e.message || e));
});
window.addEventListener("offline", () => {
  console.log("[M365OWA] network offline — renewal paused, will retry when online");
  // optimistically mark offline instantly for responsive UI; the next probe
  // will confirm or correct this.
  _netOnline = false;
  _netProbeTs = Date.now();
  _broadcastConnectivity(false);
});

// Listen for token/config changes and control messages from the options page.
browser.runtime.onMessage.addListener((msg) => {
  return (async () => {
    // dispatch based on the message type
    switch (msg && msg.type) {
      case "m365-owa-connect":
        // Start the PKCE login flow (90-day rolling refresh tokens, no Azure app registration).
        // Runs in the background — the options page polls for auth status.
        // Falls back to the OWA tab login if PKCE fails (e.g. scopes not available).
        pkceLogin().then((r) => {
          if (!r.ok) {
            console.log("[M365OWA] PKCE login failed (" + r.reason + "), falling back to OWA tab login");
            Auth.ensureOwaTab().catch((e) => console.warn("[M365OWA] OWA tab fallback failed:", e.message || e));
          }
        }).catch((e) => {
          console.warn("[M365OWA] PKCE login error:", e.message || e, "— falling back to OWA tab login");
          Auth.ensureOwaTab().catch(() => {});
        });
        // report success immediately (the login tab opens asynchronously)
        return { ok: true };
      case "m365-owa-disconnect":
        // close the OWA login tab
        await Auth.closeOwaTab();
        // clear the stored token
        await Auth.logout();
        // report success
        return { ok: true };
      case "m365-owa-relogin":
        // reload config and token after a re-login request
        await loadConfig();
        await Auth.loadStoredToken();
        // if now authenticated, re-init contacts and calendar bootstrap
        if (Auth.isAuthenticated()) {
          try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
          // re-run calendar bootstrap so the provider calendar is created now that we have a token
          try { await CalendarSync._bootstrap(); } catch (e) { console.error("[M365OWA] calendar bootstrap failed", e); }
          // trigger a calendar sync
          await messenger.calendar.calendars.synchronize().catch(() => {});
        }
        // return the new auth state
        return { ok: true, authenticated: Auth.isAuthenticated() };
      case "m365-owa-status":
        // return the full status snapshot
        return await globalThis.M365OWA.status();
      case "m365-owa-probe-status":
        // force a fresh OWA probe for explicit diagnostics (Test connection)
        await _runProbe(true);
        return await globalThis.M365OWA.status();
      case "m365-owa-set-token":
        // store a manually captured token
        await Auth.setToken(msg.token);
        return { ok: true };
      case "m365-owa-logout":
        // clear the token
        await Auth.logout();
        return { ok: true };
      case "m365-owa-save-config":
        // persist the options-page config overrides
        await browser.storage.local.set({
          owaHost: msg.owaHost,
          connectionName: msg.connectionName,
          pullDaysBack: Number(msg.pullDaysBack) | 0,
          pullDaysForward: Number(msg.pullDaysForward) | 0,
        });
        // reload CONFIG from the updated storage
        await loadConfig();
        return { ok: true };
      case "m365-owa-debug-renew":
        // force a background renewal on demand (Diagnostics button): try PKCE, then OWA refresh token, then hidden iframe, then tab fallback
        if (Auth.hasPkceRefreshToken()) { const r = await Auth.refreshViaPkce(); if (r.ok) return r; }
        if (Auth.hasRefreshToken()) { const r = await Auth.refreshViaRefreshToken(); if (r.ok) return r; }
        { const r = await renewTokenHidden(); if (r.ok) return r; return await renewTokenViaTab(); }
      case "m365-owa-sync-contacts":
        // trigger a manual contacts sync
        await ContactsSync.sync();
        return { ok: true };
      case "m365-owa-sync-calendar":
        // trigger a manual calendar sync
        await messenger.calendar.calendars.synchronize();
        return { ok: true };
      case "m365-owa-refresh-token-harvest":
        // content script captured tokens from OWA's MSAL.js cache (refresh + access)
        console.log("[M365OWA] received token harvest from " + (msg.source || "content script") +
          " (refresh_tokens=" + (msg.refresh_tokens ? msg.refresh_tokens.length : 0) +
          ", access=" + (msg.access_token ? "yes" : "no") + ")");
        {
          const wasAuth = Auth.isAuthenticated();
          let ok = false;
          if (msg.refresh_tokens && msg.refresh_tokens.length) {
            // iterate over every refresh token the content script found
            for (const t of msg.refresh_tokens) {
              const source = msg.source || "content-script";
              const r = await Auth.harvestFromTokenResponse(
                t.refreshToken, msg.access_token, t.clientId, t.url, t.scope, source
              );
              if (r) ok = true;
            }
          } else if (msg.access_token) {
            // access-token-only harvest: MSAL cache had no refresh token, but a live
            // access token is still enough to authenticate immediately. Persist it
            // directly so the addon works now; the refresh token may appear on a
            // later poll once MSAL performs a token-endpoint call.
            try { await Auth.setToken(msg.access_token); ok = true; }
            catch (e) { console.warn("[M365OWA] access-token-only harvest failed:", e.message || e); }
          }
          if (ok) {
            _clearRenewRetry();
            if (!wasAuth && Auth.isAuthenticated()) onFirstHarvest();
          }
          return { ok };
        }
    }
    // unknown message type
    return null;
  })();
});

// Debug helpers exposed on the global M365OWA object (Tools > Developer Tools > Error Console).
globalThis.M365OWA = {
  // Return a status snapshot: auth state, host, probe result, contacts AB, calendar provider, token age.
  async status() {
    // probe result placeholder
    let me = null;
    // The OWA probe (a GetCalendarView call to Microsoft) is NOT run here
    // because status() is called every 5s by the badge poll. It is only
    // run on explicit diagnostics (Test connection / Show status) and on
    // events (startup / network reconnection) via _runProbe().
    // owaProbe stays null in the polled snapshot; the badge does not use it.
    // assemble and return the status object
    return {
      authenticated: Auth.isAuthenticated(),
      // "configured" = the user has logged in before and the addon holds a
      // refresh token (PKCE 90-day or OWA 24h SPA) or a live access token. The
      // options Connect button shows "Disconnect" in this state even when the
      // access token is currently expired/pending renewal, so the user can
      // tell the account is set up rather than seeing "Connect" on every cold boot.
      configured: Auth.hasPkceRefreshToken() || Auth.hasRefreshToken() || !!Auth._token,
      // a renewal attempt (PKCE / OWA refresh / hidden iframe / tab) is running now
      renewing: _renewInProgress,
      // a backoff retry is scheduled (last renewal failed and will be retried)
      retryPending: _renewRetryTimer != null,
      // browser-reported network connectivity: probe Microsoft directly
      // (navigator.onLine is unreliable in the TB background context)
      online: (typeof navigator !== "undefined" && navigator.onLine === false) ? false : await _probeConnectivity(),
      host: CONFIG.OWA_HOST,
      connection: CONFIG.CONNECTION_NAME,
      owaProbe: _owaProbeResult,
      contactsAB: ContactsSync.abId,
      providerCal: CalendarSync.tbCalId,
      tokenAgeSec: Math.floor(Auth.tokenAgeMs() / 1000),
      tokenExpiry: Auth.getTokenExpiry(),
      hasRefreshOWA: Auth.hasRefreshToken(),
      hasRefreshPKCE: Auth.hasPkceRefreshToken(),
      owaTabOpen: (await Auth._liveOwaTabId()) != null,
    };
  },
  // Store a raw token via Auth.setToken.
  async setToken(raw) { await Auth.setToken(raw); console.log("Token stored. Reload addon or call M365OWA.reload()"); },
  // Log out by clearing the token.
  async logout() { await Auth.logout(); console.log("Logout OK"); },
  // Trigger a contacts sync.
  async syncContacts() { return ContactsSync.sync(); },
  // Trigger a calendar sync.
  async syncCalendar() { return messenger.calendar.calendars.synchronize(); },
  // Start the PKCE OAuth2 login flow (90-day rolling refresh tokens).
  async pkceLogin() { return pkceLogin(); },
  // Reload the addon.
  async reload() { browser.runtime.reload(); },
};
