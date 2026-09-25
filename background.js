// Startup entry point: loads config and token, registers the token harvester, then bootstraps contacts (best-effort) and the calendar provider.
(async function() {
  // load runtime config overrides from storage
  await loadConfig();
  // restore the stored bearer token
  await Auth.loadStoredToken();
  // register the OWA Authorization-header harvester
  registerHarvester();
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
  // opportunistically renew a stale token left over from a previous session
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
      // scan all request headers
      for (const h of details.requestHeaders) {
        // match the Authorization header case-insensitively
        if (h.name && h.name.toLowerCase() === "authorization" && h.value) {
          // remember whether we were authenticated before this harvest
          const wasAuth = Auth.isAuthenticated();
          // store the harvested token (fire-and-forget promise)
          Auth.harvestFromHeader(h.value).then((ok) => {
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

// OWA page URLs loaded into the hidden renewal frame.
function owaPageUrls() {
  // return one match pattern per first-party OWA host
  return [
    "https://outlook.office.com/owa/*",
    "https://outlook.office365.com/owa/*",
    "https://outlook.cloud.microsoft/owa/*",
  ];
}

// Strip framing protections from OWA responses while a hidden renewal runs (registered temporarily).
function stripFrameHeaders(details) {
  // collect the headers to keep
  const kept = [];
  // scan all response headers
  for (const h of details.responseHeaders || []) {
    // normalise the header name
    const name = (h.name || "").toLowerCase();
    // drop clickjacking headers so the hidden frame may load OWA
    if (name === "x-frame-options" || name === "frame-options") continue;
    // rewrite CSP without the frame-ancestors directive
    if (name === "content-security-policy") {
      // split, drop frame-ancestors, rejoin
      const v = String(h.value || "").split(";").map((s) => s.trim()).filter((s) => s && !/^frame-ancestors/i.test(s)).join("; ");
      // keep the rewritten policy when non-empty
      if (v) kept.push({ name: h.name, value: v });
      // skip the original header
      continue;
    }
    // keep everything else untouched
    kept.push({ name: h.name, value: h.value });
  }
  // return the filtered headers
  return { responseHeaders: kept };
}

// True while a hidden renewal attempt is running.
let _renewalActive = false;

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

// Renew the token invisibly: load OWA in the hidden frame so its session mints a fresh token.
async function renewTokenHidden() {
  // never overlap two renewals
  if (_renewalActive) return false;
  // never act without a previously captured token (i.e. user never connected)
  if (!Auth._token) return false;
  // mark the renewal as running
  _renewalActive = true;
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
    // log the outcome
    console.log("[M365OWA] hidden renewal " + (ok ? "succeeded" : "timed out (session may have expired)"));
    // report the outcome
    return ok;
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

// Renew once a token exists and is older than 30 minutes (even when flagged expired).
async function maybeRenew() {
  // skip when no token was ever captured
  if (!Auth._token) return;
  // skip fresh tokens
  if (Auth.tokenAgeMs() <= 30 * 60 * 1000) return;
  // run the hidden renewal; the harvester picks up the fresh token
  await renewTokenHidden();
}

// Renew the token before it goes stale via the hidden renewal frame.
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
      case "m365-owa-bookmarklet":
        // return the token-capture bookmarklet URL
        return { url: Auth.bookmarklet() };
      case "m365-owa-debug-renew":
        // force a hidden background renewal on demand (Diagnostics button)
        return { ok: await renewTokenHidden() };
      case "m365-owa-sync-contacts":
        // trigger a manual contacts sync
        await ContactsSync.sync();
        return { ok: true };
      case "m365-owa-sync-calendar":
        // trigger a manual calendar sync
        await messenger.calendar.calendars.synchronize();
        return { ok: true };
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
      owaTabOpen: (await Auth._liveOwaTabId()) != null,
    };
  },
  // Store a raw token via Auth.setToken.
  async setToken(raw) { await Auth.setToken(raw); console.log("Token stored. Reload addon or call M365OWA.reload()"); },
  // Log out by clearing the token.
  async logout() { await Auth.logout(); console.log("Logout OK"); },
  // Return the token-capture bookmarklet URL.
  bookmarklet() { return Auth.bookmarklet(); },
  // Trigger a contacts sync.
  async syncContacts() { return ContactsSync.sync(); },
  // Trigger a calendar sync.
  async syncCalendar() { return messenger.calendar.calendars.synchronize(); },
  // Reload the addon.
  async reload() { browser.runtime.reload(); },
};
