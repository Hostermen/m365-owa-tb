<h1 align="center">M365 OWA Sync (Contacts + Calendar)</h1>

<p><img src="icon.svg" width="96" height="96" alt="M365 OWA Sync"></p>

Bidirectional synchronisation of Microsoft 365 contacts and calendar
with Thunderbird, using Outlook on the web (OWA) service endpoints.

Click **Connect** in the addon settings to log in. The addon first opens
a Microsoft sign-in page and performs an OAuth2 PKCE login directly
against Microsoft's identity platform (`login.microsoftonline.com`),
using the same OAuth2 client ID Thunderbird itself uses for Microsoft
OAuth2 — no Azure app registration required. This yields a 90-day
rolling refresh token that is renewed automatically, so you log in once
and stay connected (Thunderbird only needs to run at least once every
90 days). If the PKCE login is unavailable in your tenant, the addon
falls back to opening Outlook on the web in a Thunderbird tab and
captures the session token from your own OWA traffic.

## Features

- **Calendar sync** — pull events from OWA into a Thunderbird calendar and
  push locally-created, edited, or deleted events back to OWA.
  Timezone-safe (UTC-normalised).
- **Contacts sync** — pull contacts into a Thunderbird address book and
  push contact changes back to OWA.
- **OAuth2 PKCE login (primary)** — log in once via Microsoft's identity
  platform and get 90-day rolling refresh tokens that are exchanged for
  fresh access tokens automatically — no re-login needed.
- **Automatic token refresh (fallback)** — captures fresh tokens from
  your OWA session and renews them in the background frame before
  they expire.
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
3. Open the addon settings and click **Connect** — a Microsoft sign-in
   page opens (primary PKCE path). Log in with your M365 account and
   consent to calendar/contacts access; the settings page shows
   "Connected ✓" once the token is obtained. If the PKCE login fails,
   Outlook on the web opens as a fallback and the token is captured
   after you log in there.
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
| `auth.js` | Token store, PKCE refresh, OWA-tab login, session token capture |
| `oauth.js` | OAuth2 PKCE flow against Microsoft's identity platform (Thunderbird's public client ID) |
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
- `options_ui/particles.js` — [particles.js](https://github.com/VincentGarreau/particles.js)
  by Vincent Garreau, MIT license (vendored from the official npm release 2.0.0).
