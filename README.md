<h1 align="center">M365 OWA Sync (Contacts + Calendar)</h1>

<p><img src="icon.svg" width="96" height="96" alt="M365 OWA Sync"></p>

Bidirectional synchronisation of Microsoft 365 contacts and calendar
with Thunderbird, using Outlook on the web (OWA) service endpoints.

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
   [releases](./releases) page (or build it — see below).
2. In Thunderbird: **Add-ons & Themes → gear menu → Install Add-on From
   File…** → select the `.xpi`.
3. Open the addon settings, select your M365 account from the dropdown,
   and click **Connect**.
4. Click **Sync now** to pull your contacts and calendar.

## Build

The addon is plain JavaScript with no build step. Use
[web-ext](https://extensionworkshop.com/documentation/develop/getting-started-with-web-ext/)
(Mozilla's official tool):

```sh
git clone <repo-url> m365-owa-tb
cd m365-owa-tb
npx web-ext build --overwrite-dest -a .
```

This produces `m365_owa_sync_contacts_calendar_-<version>.zip`.
Rename it to `m365-owa-tb.xpi` and install in Thunderbird.

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
