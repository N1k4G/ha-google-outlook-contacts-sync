# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
