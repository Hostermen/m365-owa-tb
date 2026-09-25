# Privacy Policy for M365 OWA Sync (Contacts + Calendar)

**Last updated: 2026-09-25**

## Summary

M365 OWA Sync synchronizes your Microsoft 365 contacts and calendar with
Thunderbird. The addon does **not** collect, store, transmit, or share your
personal data with anyone other than Microsoft (whose data it already is).

## What data the addon accesses

- **Contacts** — your Microsoft 365 contacts are fetched from OWA and stored
  in a Thunderbird address book. Edits made in Thunderbird are pushed back to
  OWA.
- **Calendar events** — your Microsoft 365 calendar events are fetched from
  OWA and stored in a Thunderbird calendar. Events created, edited, or
  deleted in Thunderbird are pushed back to OWA.
- **Session tokens** — when you click Connect, the addon opens Outlook on
  the web in a Thunderbird tab. After you log in, the addon reads the
  `Authorization` bearer token from your own OWA requests (observed via
  the `webRequest` API, never modified) and stores it in
  `browser.storage.local`. Before the token expires, the addon loads OWA
  in a hidden background frame (using your stored OWA session) so a
  fresh token is issued — no visible tab is needed. No Azure app
  registration and no Thunderbird mail account are involved.
- **Framing protection** — OWA forbids being embedded in other pages. For
  the duration of a background renewal only (max ~90 seconds), the addon
  lifts OWA's `X-Frame-Options` / `frame-ancestors` restrictions for its
  own hidden frame, then restores them immediately. Your OWA traffic is
  never modified, only observed.
- **Login page interaction** — you type your credentials directly into
  Microsoft's Outlook on the web page. The addon never sees, touches, or
  stores your password.
- **Addon settings** — sync range and connection name are stored in
  `browser.storage.local`.

## Where data goes

All data flows exclusively between your Thunderbird installation and
Microsoft's OWA servers (`outlook.office.com`, `outlook.office365.com`,
`outlook.cloud.microsoft`) over HTTPS. The addon does not communicate with
any other server. There is no analytics, telemetry, error reporting, or
tracking.

## What the addon does NOT do

- Does **not** send your data to any third party.
- Does **not** include analytics or tracking libraries.
- Does **not** upload your data to any service other than Microsoft OWA.
- Does **not** access, read, or transmit your email messages.
- Does **not** see or store your password — login happens on Microsoft's
  own page.
- Does **not** modify your OWA traffic — requests are only observed to
  read the session token.

## Third-party code

- `options_ui/particles.min.js` — a visual animation library
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
