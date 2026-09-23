<h1 align="center">M365 OWA Sync (Contacts + Calendar)
<img src="icon.svg" width="96" height="96" align="right" hspace="16" vspace="16" alt="M365 OWA Sync"></h1>

<br clear="right"><br>

Bidirectional synchronisation of Microsoft 365 contacts and calendar
with Thunderbird, using Outlook on the web (OWA) internal service
endpoints.

The addon integrates with Thunderbird's built-in OAuth2 infrastructure
to refresh access tokens automatically. When a Microsoft 365 mail
account is configured in Thunderbird with OAuth2, the addon can obtain
fresh access tokens without any external application registration or
manual credential management.

## Features

- **Calendar sync** — pull events from OWA into a Thunderbird calendar and
  push locally-created, edited, or deleted events back to OWA.
  Timezone-safe (UTC-normalised).
- **Contacts sync** — pull contacts into a Thunderbird address book and
  push contact changes back to OWA.
- **Automatic token refresh** — reuses the OAuth2 refresh tokens that
  Thunderbird already stores for the mail account. No manual token
  capture is required in normal operation.
- **Auto-fetched folder names** — calendar and contacts folder display
  names are retrieved from OWA at startup.

## Requirements

- Thunderbird 128.0 or newer (tested on 156).
- A Microsoft 365 mail account configured in Thunderbird using OAuth2
  (IMAP or EWS authentication method).

## Installation

1. Download the latest `m365-owa-tb.xpi` from the
   [releases](../../releases) page (or build it — see below).
2. In Thunderbird: **Add-ons & Themes → gear menu → Install Add-on From
   File…** → select the `.xpi`.
3. Open the addon settings, select your M365 account from the dropdown,
   and click **Connect**.
4. Click **Sync now** to pull your contacts and calendar.

## Build

The addon is plain JavaScript with no build step.

```sh
git clone <repo-url> m365-owa-tb
cd m365-owa-tb
rm -f m365-owa-tb.xpi
find . -type f -not -path "./.git/*" -not -name "*.xpi" -not -name ".DS_Store" \
  -print0 | sort -z | xargs -0 zip m365-owa-tb.xpi > /dev/null
```

Install the resulting `m365-owa-tb.xpi` in Thunderbird.

## How it works

The addon communicates with the OWA internal endpoint
(`<OWA_HOST>/owa/service.svc?action=…&app=…`), the same service endpoint
used by the Outlook web application internally. It issues EWS-shaped
SOAP operations (`FindItem`, `CreateItem`, `UpdateItem`, `DeleteItem`,
`GetFolder`, `GetCalendarView`) over this endpoint rather than the
public EWS endpoint (`/EWS/Exchange.asmx`) or EAS/ActiveSync.

### Token refresh

The experiment API in `experiments/oauth/parent/ext-oauth.js` exposes
`messenger.oauth.getAccessToken(hostname, username, type)`, which
delegates to Thunderbird's internal `OAuth2Module` to obtain fresh
access tokens using the refresh tokens Thunderbird already stores for
the configured mail account. A 5-minute timer in `auth.js` keeps the
token current; on HTTP 401/403 from OWA the addon re-fetches
immediately.

### Manual token fallback

When auto-refresh is unavailable (e.g. the account uses a non-OAuth2
authentication method), advanced settings provide a manual token
capture flow:

1. Click **Generate bookmarklet**.
2. Save the generated bookmark.
3. Open Outlook on the web in a browser, logged in.
4. Click the bookmark — it copies the bearer token to the clipboard.
5. Paste it into the addon's "Paste bearer token" field and save.

## Architecture

| File | Role |
|---|---|
| `manifest.json` | WebExtension manifest (MV2) + experiment API declarations |
| `background.js` | Lifecycle, message router, sync orchestration |
| `config.js` | Configuration load/save from `storage.local` |
| `auth.js` | Token management, auto-refresh, OAuth2 bridge |
| `owa.js` | OWA `service.svc` client (calendar + contacts operations) |
| `vcard.js` | vCard 4.0 parse/format + OWA ID mapping |
| `jcal.js` | jCal (iCalendar JSON) parse/format, UTC timezone handling |
| `contacts.js` | Contacts sync (pull/push/listen) |
| `calendar.js` | Calendar sync (pull/push via provider experiment) |
| `options_ui/` | Settings page (HTML + JS + bundled particles.js) |
| `experiments/oauth/` | Experiment API: bridge to Thunderbird's OAuth2 token store |
| `experiments/calendar/` | Experiment APIs: un-landed calendar WebExtension API |

## License

[Mozilla Public License 2.0](LICENSE).

Bundled third-party library:
- `options_ui/particles.min.js` — [particles.js](https://github.com/VincentGarreau/particles.js)
  by Vincent Garreau, MIT license.

## Limitations

- Contact push-update is implemented as delete-then-create (OWA's People
  module rejects direct `UpdateItem` on personal contacts). The OWA
  ItemId changes on every update; the addon tracks the new id via
  `vcard.js`'s `rewriteOwaId()`.
- `FindItem` on the contacts folder may return 0 contacts in some
  tenant configurations; this is under investigation.
- The OWA `service.svc` endpoint is undocumented. As long as the
  Outlook web application functions, the addon continues to work; if
  Microsoft significantly reworks the OWA backend, the addon may
  require updates.
