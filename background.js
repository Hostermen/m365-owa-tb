// Startup entry point: loads config and token, registers the token harvester, then bootstraps contacts (best-effort) and the calendar provider.
(async function() {
  // load runtime config overrides from storage
  await loadConfig();
  // restore the stored bearer token
  await Auth.loadStoredToken();
  // restore the stored refresh token (survives reboots, no SSO dependency)
  await Auth.loadStoredRefreshToken();
  // register the OWA Authorization-header harvester
  registerHarvester();
  // register the refresh-token harvester on login.microsoftonline.com
  registerRefreshTokenHarvester();
  // schedule the periodic token-renewal alarm
  await browser.alarms.create("m365-owa-renew", { periodInMinutes: 20 });
  // log the startup auth state and configured host
  console.log("[M365OWA] startup. authenticated:", Auth.isAuthenticated(), "host:", CONFIG.OWA_HOST);

  // if already authenticated, initialise contacts sync (best-effort)
  if (Auth.isAuthenticated()) {
    try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
  }
  // initialise the calendar sync provider (self-bootstraps on first pull)
  CalendarSync.init();
  // opportunistically renew an expired/stale token left over from a previous session (cold-boot recovery)
  maybeRenew().catch((e) => console.warn("[M365OWA] startup renewal failed:", e.message || e));
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
// Tries the refresh token first (no SSO, no browser, survives reboots), then the
// non-intrusive hidden iframe, then a background OWA tab, and on failure schedules
// a backoff retry so a flaky cold-boot network still recovers.
async function maybeRenew() {
  // skip when no token was ever captured AND no refresh token stored
  if (!Auth._token && !Auth.hasRefreshToken()) return;
  // skip fresh, valid tokens — nothing to do
  if (Auth._token && !Auth._expired && Auth.tokenAgeMs() <= 30 * 60 * 1000) return;

  // try the refresh token first (fastest, no SSO dependency, survives reboots)
  if (Auth.hasRefreshToken()) {
    // remember whether we were authenticated before this renewal attempt
    // (cold boot: refresh token present but no live access token → bootstrap after)
    const wasAuth = Auth.isAuthenticated();
    console.log("[M365OWA] maybeRenew: trying refresh tokens (count=" + Auth._refreshTokens.length +
      ", first url=" + (Auth._refreshTokens[0] ? Auth._refreshTokens[0].url || "none" : "none") + ")");
    const r = await Auth.refreshViaRefreshToken();
    if (r.ok) {
      _clearRenewRetry();
      // bootstrap contacts/calendar when this renewal flipped us from
      // unauthenticated → authenticated (typical cold-boot recovery)
      if (!wasAuth && Auth.isAuthenticated()) {
        onFirstHarvest().catch((e) => console.warn("[M365OWA] post-renew bootstrap failed:", e.message || e));
      }
      return;
    }
    console.log("[M365OWA] refresh token renewal failed (" + r.reason + "), falling back to hidden iframe");
  } else {
    console.log("[M365OWA] maybeRenew: no refresh token stored, trying hidden iframe");
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
  // only handle our own renewal alarm
  if (!alarm || alarm.name !== "m365-owa-renew") return;
  // run the renewal check, logging failures
  await maybeRenew().catch((e) => console.warn("[M365OWA] background renewal failed:", e.message || e));
});

// Listen for token/config changes and control messages from the options page.
browser.runtime.onMessage.addListener((msg) => {
  return (async () => {
    // dispatch based on the message type
    switch (msg && msg.type) {
      case "m365-owa-connect":
        // open the OWA login tab (or focus the existing one)
        await Auth.ensureOwaTab();
        // report success; the harvester completes the login once OWA issues a token
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
        // force a background renewal on demand (Diagnostics button): try refresh token, then hidden iframe, then tab fallback
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
    // if authenticated, probe OWA to verify the token works
    if (Auth.isAuthenticated()) { try { await OWA.probe(); me = "ok"; } catch (e) { me = "probe failed: " + e.message; } }
    // assemble and return the status object
    return {
      authenticated: Auth.isAuthenticated(),
      host: CONFIG.OWA_HOST,
      connection: CONFIG.CONNECTION_NAME,
      owaProbe: me,
      contactsAB: ContactsSync.abId,
      providerCal: CalendarSync.tbCalId,
      tokenAgeSec: Math.floor(Auth.tokenAgeMs() / 1000),
      tokenExpiry: Auth.getTokenExpiry(),
      hasRefresh: Auth.hasRefreshToken(),
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
  // Reload the addon.
  async reload() { browser.runtime.reload(); },
};
