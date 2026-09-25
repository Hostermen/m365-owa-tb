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
- **OAuth2 access tokens** — the addon obtains access tokens from
  Thunderbird's built-in OAuth2 token store via an experiment API. Tokens
  are kept in memory and refreshed as needed. They are never written to
  disk by the addon (Thunderbird manages token persistence).
- **Account configuration** — the addon stores the selected mail account
  hostname, username, and sync preferences in `browser.storage.local`.

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
- Does **not** store OAuth2 tokens outside of Thunderbird's own token store.

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
