<h1 align="center">M365 OWA Sync (Contacts + Calendar)</h1>

<p><img src="icon.svg" width="96" height="96" alt="M365 OWA Sync"></p>

Bidirectional synchronisation of Microsoft 365 contacts and calendar
with Thunderbird, using Outlook on the web (OWA) service endpoints.

Click **Connect** in the addon settings to open Outlook on the web in a
Thunderbird tab and log in with your M365 account. The addon captures
the session token from your own OWA traffic automatically and renews it
before it expires — no Azure app registration, no Thunderbird mail
account, and no manual token handling required.

## Features

- **Calendar sync** — pull events from OWA into a Thunderbird calendar and
  push locally-created, edited, or deleted events back to OWA.
  Timezone-safe (UTC-normalised).
- **Contacts sync** — pull contacts into a Thunderbird address book and
  push contact changes back to OWA.
- **Automatic token refresh** — captures fresh tokens from your OWA
  session and reloads the OWA tab before the token expires, reopening
  it silently in the background when closed. No manual token capture
  is required in normal operation.
- **Auto-fetched folder names** — calendar and contacts folder display
  names are retrieved from OWA at startup.

## Requirements

- Thunderbird 156.0 or newer.
- A Microsoft 365 account that can log in to Outlook on the web.

## Installation

1. Download the latest `m365-owa-tb.xpi` from the
   [releases](./releases) page (or build it — see below).
2. In Thunderbird: **Add-ons & Themes → gear menu → Install Add-on From
   File…** → select the `.xpi`.
3. Open the addon settings and click **Connect** — Outlook on the web
   opens in a Thunderbird tab. Log in with your M365 account; the
   settings page shows "Connected ✓" once the token is captured.
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
| `auth.js` | Token store, OWA-tab login, session token capture |
| `owa.js` | OWA `service.svc` client (calendar + contacts operations) |
| `vcard.js` | vCard 4.0 parse/format + OWA ID mapping |
| `jcal.js` | jCal (iCalendar JSON) parse/format, UTC timezone handling |
| `contacts.js` | Contacts sync (pull/push/listen) |
| `calendar.js` | Calendar sync (pull/push via provider experiment) |
| `options_ui/` | Settings page (HTML + JS + bundled particles.js) |
| `experiments/calendar/` | Experiment APIs: un-landed calendar WebExtension API (unmodified upstream drafts) |

## License

[Mozilla Public License 2.0](LICENSE).

Bundled third-party library:
- `options_ui/particles.min.js` — [particles.js](https://github.com/VincentGarreau/particles.js)
  by Vincent Garreau, MIT license.
