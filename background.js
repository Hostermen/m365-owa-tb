// Entry point. Wires config + token, then bootstraps contacts (best-effort)
// and the calendar provider (self-bootstraps on first pull).
(async function() {
  await loadConfig();
  await Auth.loadStoredToken();
  console.log("[M365OWA] startup. authenticated:", Auth.isAuthenticated(), "host:", CONFIG.OWA_HOST);

  if (Auth.isAuthenticated()) {
    try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
  }
  CalendarSync.init();
})();

// Listen for token/config changes and control messages from the options page.
browser.runtime.onMessage.addListener((msg) => {
  return (async () => {
    switch (msg && msg.type) {
      case "m365-owa-relogin":
        await loadConfig();
        await Auth.loadStoredToken();
        if (Auth.isAuthenticated()) {
          try { await ContactsSync.init(); } catch (e) { console.error("[M365OWA] contacts init failed", e); }
          // Re-run calendar bootstrap so the provider calendar is created now
          // that we have a token (startup bootstrap bailed out pre-auth).
          try { await CalendarSync._bootstrap(); } catch (e) { console.error("[M365OWA] calendar bootstrap failed", e); }
          await messenger.calendar.calendars.synchronize().catch(() => {});
        }
        return { ok: true, authenticated: Auth.isAuthenticated() };
      case "m365-owa-status":
        return await globalThis.M365OWA.status();
      case "m365-owa-set-token":
        await Auth.setToken(msg.token);
        return { ok: true };
      case "m365-owa-logout":
        await Auth.logout();
        return { ok: true };
      case "m365-owa-configure-auto-refresh":
        await Auth.configureAutoRefresh(msg.hostname, msg.username, msg.accountType);
        return { ok: true };
      case "m365-owa-save-config":
        await browser.storage.local.set({
          owaHost: msg.owaHost,
          connectionName: msg.connectionName,
          pullDaysBack: Number(msg.pullDaysBack) | 0,
          pullDaysForward: Number(msg.pullDaysForward) | 0,
        });
        await loadConfig();
        return { ok: true };
      case "m365-owa-bookmarklet":
        return { url: Auth.bookmarklet() };
      case "m365-owa-sync-contacts":
        await ContactsSync.sync();
        return { ok: true };
      case "m365-owa-sync-calendar":
        await messenger.calendar.calendars.synchronize();
        return { ok: true };
    }
    return null;
  })();
});

// Debug helpers (Tools > Developer Tools > Error Console)
globalThis.M365OWA = {
  async status() {
    let me = null;
    if (Auth.isAuthenticated()) { try { await OWA.probe(); me = "ok"; } catch (e) { me = "probe failed: " + e.message; } }
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
  async setToken(raw) { await Auth.setToken(raw); console.log("Token stored. Reload addon or call M365OWA.reload()"); },
  async logout() { await Auth.logout(); console.log("Logout OK"); },
  bookmarklet() { return Auth.bookmarklet(); },
  async syncContacts() { return ContactsSync.sync(); },
  async syncCalendar() { return messenger.calendar.calendars.synchronize(); },
  async reload() { browser.runtime.reload(); },
};
