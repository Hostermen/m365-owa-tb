// Startup entry point: loads config and token, then bootstraps contacts (best-effort) and the calendar provider.
(async function() {
  // load runtime config overrides from storage
  await loadConfig();
  // restore the stored bearer token (and auto-refresh settings)
  await Auth.loadStoredToken();
  // log the startup auth state and configured host
  console.log("[M365OWA] startup. authenticated:", Auth.isAuthenticated(), "host:", CONFIG.OWA_HOST);

  // if already authenticated, initialise contacts sync (best-effort)
  if (Auth.isAuthenticated()) {
    try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
  }
  // initialise the calendar sync provider (self-bootstraps on first pull)
  CalendarSync.init();
})();

// Listen for token/config changes and control messages from the options page.
browser.runtime.onMessage.addListener((msg) => {
  return (async () => {
    // dispatch based on the message type
    switch (msg && msg.type) {
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
        // clear the token and stop refresh
        await Auth.logout();
        return { ok: true };
      case "m365-owa-configure-auto-refresh":
        // configure OAuth2 auto-refresh
        await Auth.configureAutoRefresh(msg.hostname, msg.username, msg.accountType);
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
  // Return a status snapshot: auth state, host, probe result, contacts AB, calendar provider, auto-refresh config.
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
      autoRefresh: Auth.isAutoRefreshEnabled(),
      oauthConfig: Auth.getOAuthConfig(),
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
