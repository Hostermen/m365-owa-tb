# Privacy Policy for M365 OWA Sync (Contacts + Calendar)

**Last updated: 2026-11-03**

## Summary

M365 OWA Sync synchronizes your Microsoft 365 contacts and calendar with
Thunderbird. The addon does **not** collect analytics, does **not** talk to
any server other than Microsoft's, and does
**not** share your personal data with any third party.

## How the addon authenticates (in order of attempt)

1. **Primary: OAuth2 PKCE login against Microsoft's identity platform.**
   When you click **Connect**, the addon first opens a Microsoft sign-in
   page (`login.microsoftonline.com`) and
   performs an OAuth2 Authorization Code flow with PKCE (Proof Key for
   Code Exchange, RFC 7636) directly against Microsoft's identity
   platform. It requests consent for calendar and contacts access
   (`Calendars.ReadWrite`, `Contacts.ReadWrite`, `offline_access`)
   This produces a short-lived access token
   (~60–90 minutes) and a long-lived refresh token (90-day rolling
   expiry). The refresh token is stored in `browser.storage.local` and
   is exchanged for new access tokens automatically, without any login
   prompt, for as long as Thunderbird runs at least once every 90 days.
2. **Fallback: OWA session-token capture.** Only if the PKCE login fails
   (for example, the requested
   scopes are unavailable), the addon falls back to opening Outlook on
   the web (OWA) in a Thunderbird tab. After you log in there, the addon
   reads the `Authorization` bearer token from your own OWA requests
   (observed via the `webRequest` API, never modified) and stores it in
   `browser.storage.local`.

## What data the addon accesses

- **Contacts** — your Microsoft 365 contacts are fetched from OWA and stored
  in a Thunderbird address book. Edits made in Thunderbird are pushed back to
  OWA.
- **Calendar events** — your Microsoft 365 calendar events are fetched from
  OWA and stored in a Thunderbird calendar. Events created, edited, or
  deleted in Thunderbird are pushed back to OWA.
- **OAuth tokens** — the access token and the 90-day rolling refresh token
  obtained from the PKCE login are stored in `browser.storage.local`. The
  addon never sees, touches, or stores your password in any flow: on the
  PKCE path you type your credentials directly into Microsoft's own
  sign-in page; on the fallback path into Microsoft's OWA page.
- **Session tokens (fallback path)** — if the PKCE login fails and the
  OWA fallback is used, the addon reads the `Authorization` bearer token
  from your own OWA requests (observed via the `webRequest` API, never
  modified) and stores it in `browser.storage.local`. The token persists
  across Thunderbird restarts and reboots. When the token expires, the
  addon renews it automatically: first by loading OWA in a hidden
  background frame; if that fails (e.g. after a cold boot or crash) by opening a
  background OWA tab using your stored session cookies. Both paths reuse
  your existing OWA session (no new login) and close/release the
  frame/tab immediately once a fresh token is captured. Failed renewals
  are retried with exponential backoff.
- **Framing Protection (fallback path)** — OWA forbids being embedded in
  other pages. For the duration of a hidden-frame renewal only (max ~90
  seconds), the addon lifts OWA's `X-Frame-Options` / `frame-ancestors`
  restrictions for its own hidden frame, then restores them immediately.
  The background-tab fallback does not alter any headers. Nothing else
  about your OWA traffic is modified — requests are only observed to
  read the session token.
- **Login page interaction** — you type your credentials directly into
  Microsoft's own pages (the Microsoft sign-in page for the PKCE flow,
  or Outlook on the web for the fallback flow). The addon never sees,
  touches, or stores any password.
- **Cookies** — logging in stores Microsoft's own session cookies in
  Thunderbird (as any Microsoft login does). The addon installs no
  cookies of its own. Cookies are required for the addon to function:
  connecting (clicking Connect) is the explicit opt-in, and Disconnect
  clears the token. Persistent session cookies (e.g.
  `ESTSAUTHPERSISTENT`) allow the addon to silently renew tokens after a
  reboot without asking you to log in again. To remove the session
  cookies, log out of OWA or clear Thunderbird's cookies.
- **Addon settings** — sync range and connection name are stored in
  `browser.storage.local`.

## Where data goes

All data flows exclusively between your Thunderbird installation and
Microsoft's servers — the Microsoft identity platform
(`login.microsoftonline.com`) for the OAuth2 login and token refresh, and
Microsoft's OWA servers (`outlook.office.com`, `outlook.office365.com`,
`outlook.cloud.microsoft`) for calendar and contacts data — over HTTPS.
The addon does not communicate with any other server. There is no
analytics, telemetry, error reporting, or tracking.

## What the addon does NOT do

- Does **not** send your data to any third party.
- Does **not** include analytics or tracking libraries.
- Does **not** upload your data to any service other than Microsoft.
- Does **not** access, read, or transmit your email messages.
- Does **not** see or store your password — login happens on Microsoft's
  own page.
- Does **not** modify your OWA requests or their content — the only
  responses ever altered are OWA's framing protections during a
  background renewal (see above), which are restored immediately.

## Third-party code

- `options_ui/particles.js` — a visual animation library
  ([particles.js](https://github.com/VincentGarreau/particles.js), MIT
  license) used solely for the settings page background. It does not
  access, collect, or transmit any data.

## Data deletion

Uninstalling the addon removes the addon code and its settings from
`browser.storage.local`. Contacts and calendar events already synced into
Thunderbird remain in Thunderbird until you delete them. Items on
Microsoft's servers are never deleted by the addon itself (only when you
explicitly delete an event or contact in Thunderbird and the change is
pushed to OWA).

## Open source

The addon is open source under the Mozilla Public License 2.0. You can
audit every line of code at
https://github.com/Hostermen/m365-owa-tb.

## Contact

For privacy questions, open an issue at
https://github.com/Hostermen/m365-owa-tb/issues.
