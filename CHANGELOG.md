# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Upgrade notes
- **Minimum Home Assistant version is now 2025.8.0.** The options flow now
  uses `OptionsFlowWithReload` instead of a config entry update listener. HA
  2026.12 turns an update listener combined with the reauth/reconfigure reload
  helper into an error. HACS will not offer this update to installations below
  2025.8.0; if you installed manually, check your version before updating.

### Added
- **Reconfigure flow for app credentials** (#36). *Settings → Devices &
  Services → Google Outlook Contacts Sync → Reconfigure* changes the Microsoft
  and Google client IDs and secrets without removing the integration, so
  tokens, the contact mapping and the sync token are kept. A secret left empty
  stays unchanged. Before anything is saved, a new secret is validated by
  redeeming the stored refresh token with it: a secret the provider rejects is
  reported in the form and nothing is written; a refresh token that can no
  longer be redeemed, or a changed client ID, leads into that provider's
  sign-in.
- README section *Rotating or changing client secrets*.

### Fixed
- **A revoked or expired Microsoft sign-in now starts re-authorization**
  (#38). When O365's token refresh failed, it raised a bare `RuntimeError`, so
  the run only counted failed contacts and never asked to re-authorize. Every
  run (sync, full resync and dry run) now checks the Microsoft token first and
  refreshes it itself when it is about to expire. A rejected client secret or
  refresh token starts the Microsoft re-authorization, where a new secret can
  be entered. An unreachable token endpoint stays an ordinary sync error.
  A refresh that fails in the middle of a run is classified the same way.
- **The Microsoft sign-in no longer lapses when no contact changes** (#38). A
  run without changes sent no Graph request, so the refresh token was never
  redeemed and expired after Microsoft's inactivity window (90 days by
  default). The token check at the start of every run now keeps it in use.
- **Re-authorization no longer loops on a revoked or expired client secret**
  (#36). The re-authentication prompt now accepts a new secret for the affected
  provider (empty keeps the stored one). Previously it always reused the stored
  secret, so the code exchange failed again, and the only way out was deleting
  the integration or editing `.storage/core.config_entries` by hand.
- Re-authorization replaced the whole entry data and dropped keys it does not
  manage, such as a legacy `delete_removed`; it now only writes what it
  changed.
- A sync still running when a reauth or reconfigure finished could write its
  old token back over the new one before the reload, undoing the new sign-in.
  Tokens replaced from outside a run are now left alone.
- After a failed sign-in during setup, reauth or reconfigure, the retry form
  keeps the client ID, and client secrets are entered in masked fields.
- The manifest now accepts `google-api-python-client` 2.71.0, the version
  Home Assistant pins for its own Google integrations. The previous
  `>=2.100.0` failed hassfest and could force an upgrade under them.

## [1.1.0] - 2026-07-27

### Upgrade notes
- **Minimum Home Assistant version is now 2024.12.0.** `hacs.json` previously
  allowed 2024.1.0, on which the integration cannot import. HACS will not offer
  this update to installations below 2024.12.0; if you installed manually, check
  your version before updating.
- **Every mapped contact is updated once on the first run after upgrading.**
  Change detection now hashes the full authoritative field set rather than a
  sparse payload, so all stored hashes are invalidated. Expected, one-time, and
  self-limiting.
- **"Synced contacts" may drop.** Contacts deleted in Google are now untracked
  even when *Delete removed contacts* is off; previously they were counted
  forever. The lower number is the accurate one.
- **The `sync_now` service is debounced.** It now uses the same refresh path as
  the Sync now button instead of forcing an immediate run, so rapid repeat calls
  collapse into a single sync. Automations that assumed the sync had already
  started when the service call returned should await the result via the
  Last sync sensor instead.

### Fixed
- Contact creation failed on every run: the Graph id was read off the
  `requests.Response` object instead of its JSON body, so no contact was ever
  created and the Google sync token never advanced (#6).
- Failed Graph reads returned partial or empty contact lists as success, which
  could drive adoption and duplicate removal from an incomplete picture (#26).
- Failed Graph writes were counted as successes, storing mappings with an empty
  contact id (#10).
- Fields removed in Google were never cleared in Outlook, and the run was
  recorded as synced anyway — a permanent silent divergence (#24).
- Google addresses of type `other` or an unknown type were stored as business
  addresses; `otherAddress` is now used (#32).
- Google email *types* were written into the Graph `name` field, which is a
  display name (#29).
- Full resync adopted Outlook contacts by name alone and could overwrite an
  unrelated person with the same name. Adoption now requires exactly one
  unambiguous email match (#25).
- Full resync could not recover a mapped Outlook contact deleted by hand (#30).
- Mappings for contacts deleted in Google were never cleaned up (#18).
- Duplicate removal could merge two tracked Google contacts onto one Outlook
  contact, permanently corrupting the mapping (#31).
- Duplicate removal swallowed Microsoft auth errors and did not mark partial
  failures, leaving the problem sensor off (#27).
- A missing Google token crashed setup instead of offering re-authorization (#16).
- Refreshed OAuth tokens were never persisted and were lost on restart (#7).
- Overlapping sync, full resync and dry-run passes could interleave and create
  duplicates (#8).
- A missing or expired sync token ran the delta pipeline over a full fetch
  rather than a real reconciliation (#33).
- A failed Google fetch was reported as a successful run and advanced the
  last-sync sensor (#14).
- Setup blocked on a full network sync; the initial run is now a background
  task (#19).
- Services were never removed on unload, and OAuth callback views were
  registered repeatedly (#17, #20).
- Duplicate-detection scoring counted fields that were never requested from
  Graph, so a data-rich contact could be deleted in favour of an emptier
  one (#9).
- OAuth error descriptions were reflected into an unauthenticated HTML
  response without escaping (#13).
- `ContactSyncStore` instances shared mutable default containers, leaking
  mappings between instances in one process (#28).

### Changed
- A failed full resync now clears the sync token so the next run reconciles in
  full again, instead of resuming from a stale delta cursor.
- Full resync adopts an existing Outlook contact only on a unique email match,
  and auto-dedup skips groups containing more than one tracked contact.
- Documentation corrected: HA's `.storage/` is **not** encrypted, and the Google
  step needs a **Web application** OAuth client, not a Desktop client (#12, #15).
- Added a trademark and API-compliance disclaimer (#23).

## [1.0.0] - 2026-06-16

### Added
- Initial release.
- One-way sync from Google Contacts (People API) to Microsoft Outlook (Graph API) via direct Microsoft Graph REST calls.
- Fields synced: name, emails, phone numbers, birthdays (including year-less contacts using the 1604 Outlook placeholder), addresses, company, job title, department.
- Delta sync via Google `syncToken` — only changed contacts are transferred after the first full run. Sync token is held back when any contact fails, so failures are retried on the next run.
- Full resync mode (`full_sync` service / **Full resync** button): fetches all Google contacts, adopts existing Outlook contacts by name to avoid duplicates, pushes any differences, and resets the delta token.
- Dry-run preview (**Run dry-run** button / **Dry-run preview** sensor): computes what the next delta sync would create/update/delete without applying anything.
- Duplicate detection: every sync can scan Outlook for contacts with the same normalized name and delete all but one per group (opt-in, default off). Keeps the already-mapped contact; falls back to most-complete then most-recently-modified.
- Entities grouped under one HA device:
  - Sensors: Last sync, Last sync result (with per-contact statistics), Next sync, Synced contacts, Duplicates removed, Dry-run preview.
  - Buttons: Sync now, Full resync, Run dry-run.
  - Binary sensor: Sync problem (on when the last run failed or had per-contact errors).
  - Diagnostics: redacted config entry data downloadable from the HA UI.
- Services: `sync_now` (delta sync) and `full_sync` (full reconciliation).
- Configurable options: sync interval (1–168 h, default 24 h), delete-removed toggle, auto-remove-duplicates toggle.
- Microsoft OAuth via python-o365 "alternate auth" redirect flow with a HA HTTP callback view.
- Google OAuth via authorization-code redirect with a HA HTTP callback view.
- Persistent contact ID mapping and sync token in HA's encrypted `.storage/`.
- German and English UI translations.
- HACS-compatible manifest and CI pipelines (hassfest, HACS validation, pytest, ruff, mypy).
