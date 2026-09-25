// Configuration for the OWA-backed sync.
var CONFIG = {
  // OWA host for this business account; the bearer token must have been captured from the same host.
  OWA_HOST: "https://outlook.office.com",
  // Optional human-readable label; disambiguates storage keys when multiple accounts are configured.
  CONNECTION_NAME: "default",
  // Days back from now to pull on each calendar sync.
  PULL_DAYS_BACK: 30,
  // Days forward from now to pull on each calendar sync.
  PULL_DAYS_FORWARD: 90,
  // Interval between automatic calendar syncs (10 minutes in ms).
  SYNC_INTERVAL_MS: 10 * 60 * 1000,
  // OWA service.svc actions used by this addon (see owa.js for payloads).
  // Calendar: GetFolder, GetCalendarView, CreateItem, UpdateItem, DeleteItem
  // People  : GetFolder, FindPeople, FindItem, CreateItem, UpdateItem, DeleteItem
};

// Load runtime overrides from storage.local (written by the options page) and apply them to CONFIG.
async function loadConfig() {
  // batch-read all override keys from storage
  const s = await browser.storage.local.get([
    "owaHost",
    "connectionName",
    "pullDaysBack",
    "pullDaysForward",
  ]);
  // if a host override exists
  if (s.owaHost) {
    try {
      // parse it as a URL and keep only the origin (scheme + host + port)
      CONFIG.OWA_HOST = new URL(s.owaHost).origin;
    } catch {
      // if URL parsing fails, strip trailing slashes and any /owa/ path suffix
      CONFIG.OWA_HOST = String(s.owaHost).replace(/\/+$/, "").replace(/\/owa\/.*$/i, "");
    }
  }
  // override connection name if present
  if (s.connectionName) CONFIG.CONNECTION_NAME = String(s.connectionName);
  // override pull-days-back if present
  if (s.pullDaysBack) CONFIG.PULL_DAYS_BACK = Number(s.pullDaysBack) | 0;
  // override pull-days-forward if present
  if (s.pullDaysForward) CONFIG.PULL_DAYS_FORWARD = Number(s.pullDaysForward) | 0;
  console.log("[M365OWA] config loaded. host:", CONFIG.OWA_HOST, "connection:", CONFIG.CONNECTION_NAME);
}
