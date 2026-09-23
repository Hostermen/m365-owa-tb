# M365 OWA Sync (Contacts + Calendar)

Bidirectional sync of Microsoft 365 contacts and calendar into Thunderbird,
using Outlook on the web (OWA) internal service endpoints.

No Azure app registration required. The addon piggybacks on Thunderbird's
built-in OAuth2 to refresh tokens automatically — as long as you have an
IMAP/OWA mail account set up in Thunderbird with OAuth2, the addon can fetch
fresh access tokens on its own.

## Features

- **Calendar sync** — pull events from OWA into a Thunderbird calendar, push
  locally-created/edited/deleted events back to OWA. Timezone-safe (UTC).
- **Contacts sync** — pull contacts into a Thunderbird address book, push
  contact changes back to OWA (delete + create on update).
- **Auto token refresh** — reuses Thunderbird's stored OAuth2 refresh tokens
  via an experiment API; no manual token capture needed in normal use.
- **Auto-fetched folder names** — calendar and contacts folder display
  names are pulled from OWA (e.g. "Kalender", "Kontakte").
- **Contact event listeners** — local contact edits trigger sync with 2s
  debounce for near-real-time push.
- **Clean settings UI** — green theme with animated particle background;
  account picker, test, and sync in the main view; advanced config
  (sync range, manual token fallback, diagnostics) collapsible.

## Requirements

- Thunderbird 128.0 or newer (tested on 156).
- A Microsoft 365 mail account configured in Thunderbird using OAuth2
  (IMAP or EWS auth method).

## Install

1. Download the latest `m365-owa-tb.xpi` from the [releases](../../releases) page
   (or build it — see below).
2. In Thunderbird: **Add-ons & Themes → gear menu → Install Add-on From
   File…** → select the `.xpi`.
3. Open the addon settings, pick your M365 account from the dropdown, and
   click **Connect**.
4. Click **Sync now** to pull your contacts and calendar.

## Build

The addon is plain JavaScript — no build step is required.

```sh
git clone <repo-url> m365-owa-tb
cd m365-owa-tb
rm -f m365-owa-tb.xpi
find . -type f -not -path "./.git/*" -not -name "*.xpi" -not -name ".DS_Store" \
  -print0 | sort -z | xargs -0 zip m365-owa-tb.xpi > /dev/null
```

Install the resulting `m365-owa-tb.xpi` in Thunderbird.

## How it works

The addon talks to the OWA internal endpoint
(`<OWA_HOST>/owa/service.svc?action=…&app=…`) — the same one the Outlook
web app uses internally. It speaks EWS-shaped SOAP operations
(`FindItem`, `CreateItem`, `UpdateItem`, `DeleteItem`, `GetFolder`,
`GetCalendarView`) but over the OWA service endpoint, **not** the public
EWS endpoint (`/EWS/Exchange.asmx`) or EAS/ActiveSync.

### Token refresh

`experiments/oauth/parent/ext-oauth.js` exposes
`messenger.oauth.getAccessToken(hostname, username, type)` which calls
Thunderbird's internal `OAuth2Module` to obtain fresh access tokens using
the refresh tokens Thunderbird already stores for your mail account. A
5-minute timer in `auth.js` keeps the token fresh; on 401/403 from OWA the
addon re-fetches immediately.

### Manual token fallback

If auto-refresh is unavailable (e.g. your account uses a non-OAuth2 auth
method), the advanced settings expose a **manual token capture** flow:

1. Click **Generate bookmarklet**.
2. Save the generated bookmark to your bookmarks.
3. Open Outlook on the web in a browser, logged in.
4. Click the bookmark — it copies your bearer token to the clipboard.
5. Paste it into the addon's "Paste bearer token" field and save.

## Architecture

| File | Role |
|---|---|
| `manifest.json` | WebExtension manifest (MV2) + experiment API declarations |
| `background.js` | Lifecycle, message router, sync orchestration |
| `config.js` | Config load/save from `storage.local` |
| `auth.js` | Token management, auto-refresh, OAuth2 bridge |
| `owa.js` | OWA `service.svc` client (calendar + contacts operations) |
| `vcard.js` | vCard 4.0 parse/format + OWA ID mapping |
| `jcal.js` | jCal (iCalendar JSON) parse/format, UTC timezone handling |
| `contacts.js` | Contacts sync (pull/push/listen) |
| `calendar.js` | Calendar sync (pull/push via provider experiment) |
| `options_ui/` | Settings page (HTML + JS + bundled particles.js) |
| `experiments/oauth/` | Experiment API: bridge to TB's OAuth2 token store |
| `experiments/calendar/` | Experiment APIs: un-landed calendar WebExtension API |

## License

[Mozilla Public License 2.0](LICENSE).

Bundled third-party library:
- `options_ui/particles.min.js` — [particles.js](https://github.com/VincentGarreau/particles.js)
  by Vincent Garreau, MIT license.

## Limitations

- Contact push-update is implemented as delete-then-create (OWA's People
  module rejects direct `UpdateItem` on personal contacts). This means the
  OWA ItemId changes on every update — the addon tracks the new id via
  `vcard.js`'s `rewriteOwaId()`.
- `FindItem` on the contacts folder may return 0 contacts in some tenant
  configurations; this is under investigation.
- The OWA `service.svc` endpoint is undocumented. As long as the Outlook
  web app works, the addon works; if Microsoft significantly reworks OWA's
  backend, the addon may break.
