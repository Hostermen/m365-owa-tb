// Configuration for the OWA-backed sync.
//
// Unlike the Graph variant, this addon does NOT use OAuth / an Azure app
// registration. Instead it reuses the strategy from the `m365-owa-cli` project:
// the user captures an OWA bearer token from a logged-in Outlook on the web
// browser tab (via a bookmarklet or DevTools) and pastes it into the addon
// options page. That token is sent to OWA's internal service endpoint
// (POST <OWA_HOST>/owa/service.svc?action=...&app=...).
//
// All settings can be overridden at runtime from storage.local (set via the
// options page). Nothing is hardcoded per-tenant.
var CONFIG = {
  // Where Outlook on the web lives for this business account. The bearer token
  // must have been captured from the SAME host. Common values:
  //   https://outlook.office.com
  //   https://outlook.office365.com
  //   https://outlook.cloud.microsoft
  OWA_HOST: "https://outlook.office.com",

  // Optional human label for the captured connection (mirrors m365-owa-cli
  // --connection). Purely informational.
  CONNECTION_NAME: "default",

  // Default calendar window pulled on each sync: days forward + days back from now.
  PULL_DAYS_BACK: 30,
  PULL_DAYS_FORWARD: 90,

  // Names are fetched from OWA at runtime (see owa.js getCalendarFolderName /
  // getContactsFolderName). These are only fallbacks if the OWA fetch fails.
  AB_NAME: "M365 OWA Contacts",
  CAL_NAME: "M365 OWA Calendar",
  SYNC_INTERVAL_MS: 10 * 60 * 1000,

  // OWA service.svc actions (kept here for visibility; see owa.js for payloads).
  // Calendar: GetCalendarFolders, GetCalendarView, CreateItem, DeleteItem
  // People  : FindPeople            (best-effort; m365-owa-cli has no contacts)
};

// Load runtime overrides from storage.local. The options page writes:
//   { owaHost, connectionName, pullDaysBack, pullDaysForward, abName, calName }
async function loadConfig() {
  const s = await browser.storage.local.get([
    "owaHost",
    "connectionName",
    "pullDaysBack",
    "pullDaysForward",
  ]);
  if (s.owaHost) {
    try {
      CONFIG.OWA_HOST = new URL(s.owaHost).origin;
    } catch {
      CONFIG.OWA_HOST = String(s.owaHost).replace(/\/+$/, "").replace(/\/owa\/.*$/i, "");
    }
  }
  if (s.connectionName) CONFIG.CONNECTION_NAME = String(s.connectionName);
  if (s.pullDaysBack) CONFIG.PULL_DAYS_BACK = Number(s.pullDaysBack) | 0;
  if (s.pullDaysForward) CONFIG.PULL_DAYS_FORWARD = Number(s.pullDaysForward) | 0;
  console.log("[M365OWA] config loaded. host:", CONFIG.OWA_HOST, "connection:", CONFIG.CONNECTION_NAME);
}
