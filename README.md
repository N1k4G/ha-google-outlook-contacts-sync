# Google Outlook Contacts Sync — Home Assistant Integration

[![Tests](https://github.com/N1k4G/ha-google-outlook-contacts-sync/actions/workflows/test.yaml/badge.svg)](https://github.com/N1k4G/ha-google-outlook-contacts-sync/actions/workflows/test.yaml)
[![Hassfest](https://github.com/N1k4G/ha-google-outlook-contacts-sync/actions/workflows/hassfest.yaml/badge.svg)](https://github.com/N1k4G/ha-google-outlook-contacts-sync/actions/workflows/hassfest.yaml)
[![HACS](https://img.shields.io/badge/HACS-Custom-orange.svg)](https://hacs.xyz)

One-way sync from **Google Contacts → Microsoft Outlook** (personal Microsoft account) as a Home Assistant custom integration. Runs on schedule (default every 24 h) or on demand via a service call.

> **Sync direction:** Google → Outlook only. Changes in Outlook are not reflected back.

---

## Features

- Delta sync via Google People API `syncToken` — only changed contacts are transferred after the first full run.
- Birthdays synced as a contact field; contacts without a year use the 1604 placeholder (Outlook's own convention, shown as day/month only).
- Configurable sync interval and optional deletion of Outlook contacts when removed from Google.
- Persists the contact ID mapping and sync token in HA's `.storage/` directory — survives restarts without duplication.
- `sync_now` service for immediate sync from automations or the Developer Tools.
- Sensors and buttons for monitoring and control (last sync, last result, next sync, synced count, sync-problem alert, sync-now button) plus an optional dry-run preview.
- German and English UI.

---

## Prerequisites

You need two sets of OAuth credentials — one for Microsoft, one for Google. Both are free.

### 1. Microsoft Entra App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Entra ID → App registrations → New registration**.
2. Name: anything (e.g. "HA Google Outlook Contacts Sync").
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
4. Redirect URI (Web): `https://<your-ha-external-url>/api/google_outlook_contacts_sync/auth`
5. After creation go to **API permissions → Add a permission → Microsoft Graph → Delegated** and add:
   - `Contacts.ReadWrite`
   - `User.Read`
6. Go to **Certificates & secrets → New client secret**. Copy the value immediately — it is shown only once.
7. Note the **Application (client) ID** from the Overview page.

### 2. Google Cloud Project and OAuth Client

1. Go to [console.cloud.google.com](https://console.cloud.google.com) → create a new project.
2. Enable the **People API** (APIs & Services → Library → search "People API").
3. Go to **APIs & Services → OAuth consent screen** — choose External, fill in the app name, and add your Google account as a test user.
4. Go to **APIs & Services → Credentials → Create credentials → OAuth client ID**.
5. Application type: **Web application**. Name it anything.
6. Under **Authorized redirect URIs**, add: `https://<your-ha-external-url>/api/google_outlook_contacts_sync/google_auth`
7. Note the **Client ID** and **Client secret**.

---

## Installation

**Requires Home Assistant 2024.12.0 or newer.** The integration uses
`ConfigFlowResult` (2024.4), the `config_entry` argument to
`DataUpdateCoordinator` (2024.11/12) and the `single_config_entry` manifest flag
(2024.11); on older versions setup fails at import.

### HACS (custom repository)

1. In HACS, go to **Integrations → ⋮ → Custom repositories**.
2. Add `https://github.com/N1k4G/ha-google-outlook-contacts-sync` with category **Integration**.
3. Install "Google Outlook Contacts Sync".
4. Restart Home Assistant.

### Manual

Copy the `custom_components/google_outlook_contacts_sync` folder into your HA `config/custom_components/` directory and restart.

---

## Setup

Before starting, make sure Home Assistant knows its external URL: **Settings → System → Network → Home Assistant URL**.

1. Go to **Settings → Devices & Services → Add Integration** → search "Google Outlook Contacts Sync".
2. **Step 1 — Microsoft credentials:** Enter the Client ID and Client Secret from your Entra app.
3. **Step 2 — Microsoft authorization:** A browser window opens. Sign in with your personal Microsoft account and grant contacts access. You are redirected back to HA automatically.
4. **Step 3 — Google credentials:** Enter the Client ID and Client Secret from your Google Web OAuth client.
5. **Step 4 — Google authorization:** A browser window opens. Sign in with your Google account and grant contacts access. You are redirected back to HA automatically.
6. The first full sync runs immediately. Subsequent runs use delta sync.

---

## Options

After setup go to the integration's **Configure** button:

| Option | Default | Description |
|---|---|---|
| Sync interval (hours) | 24 | How often to sync automatically (1–168 h). |
| Delete removed contacts | Off | When enabled, contacts deleted from Google are also deleted in Outlook. |
| Auto-remove duplicate contacts | Off | **Destructive.** When enabled, every sync scans Outlook for duplicate contacts (**matched by name only**) and deletes all but one per group. See [Duplicate handling](#duplicate-handling) before enabling. |

---

## Entities

All entities are grouped under one device, **Google Outlook Contacts Sync**.

| Entity | Type | Description |
|---|---|---|
| Last sync | Sensor (timestamp) | When the last **successful** sync finished. |
| Last sync result | Sensor | Short summary (e.g. `5 created, 2 updated`). Attributes: `created`, `updated`, `deleted`, `skipped`, `failed`, `duplicates_removed`, `duration_seconds`, `errors`. |
| Next sync | Sensor (timestamp) | Estimated time of the next scheduled run. |
| Synced contacts | Sensor | Number of contacts currently tracked/synced to Outlook. |
| Duplicates removed | Sensor | How many duplicate Outlook contacts the last sync deleted. Attribute `groups` lists the affected names. |
| Sync problem | Binary sensor (problem) | On when the last run failed or reported per-contact failures. Attributes carry the error details. |
| Sync now | Button | Triggers an immediate delta sync (same as the `sync_now` service). |
| Full resync | Button | Runs a full reconciliation (same as the `full_sync` service) — see below. |
| Run dry-run | Button | Recomputes the dry-run preview. |
| Dry-run preview | Sensor | Number of planned changes. Attributes `to_create` / `to_update` / `to_delete` list the affected contact names. Reflects exactly what the next delta sync would do and applies nothing. Starts empty until you press **Run dry-run**. |

---

## How change detection works

Scheduled runs use **delta sync**: Google's People API returns only the contacts that changed since the last successful sync (tracked by a `syncToken`), and the integration pushes a contact to Outlook only when its mapped fields actually differ. So a run reporting **0 created/updated/deleted is normal when nothing changed** in Google since the previous run.

If Outlook drifts out of sync — for example contacts are missing, or a past run failed — use **Full resync**.

### Full resync

The **Full resync** button (and `full_sync` service) fetches **all** Google contacts, then:

- **Adopts** an existing Outlook contact instead of creating a new one when exactly one not-yet-tracked Outlook contact shares an email address with the Google contact. Matching is on the normalized address only — a matching *name* is never enough, since two different people commonly share one.
- **Repairs** mappings whose Outlook contact was deleted by hand, recreating or re-adopting the contact.
- **Reconciles** contacts removed from Google: deleting them in Outlook when *Delete removed contacts* is on, otherwise just untracking them.
- Pushes any field differences and recreates anything genuinely missing.
- Resets the delta token so subsequent scheduled runs are clean.

Run this once after upgrading, or whenever the contact counts look wrong.

### Duplicate handling

With **Auto-remove duplicate contacts** enabled, each sync lists all Outlook contacts, groups them by **normalized name**, and deletes all but one contact per group (keeping the one already tracked, else the most complete / most recently modified). Repointed mappings keep the kept contact linked to Google.

A group containing **more than one contact already tracked by this integration is skipped entirely** and logged as a warning. Merging those would leave several Google contacts pointing at a single Outlook contact, which then overwrite each other on every sync. Auto-dedup therefore only removes untracked duplicates around at most one tracked contact, or deduplicates groups with no tracked contacts at all.

> ⚠️ **Matching is name-only and removal is automatic.** Two genuinely different people who share a name will be treated as duplicates and one copy will be deleted on every sync. The dry-run preview does **not** cover dedup, so leave this off unless you understand the risk — test on a non-critical account if unsure. Every deletion is logged at WARNING level and counted by the **Duplicates removed** sensor.

---

## Services

### `google_outlook_contacts_sync.sync_now`

Triggers an immediate delta sync. No parameters required.

```yaml
service: google_outlook_contacts_sync.sync_now
```

### `google_outlook_contacts_sync.full_sync`

Runs a full reconciliation (adopt on a unique email match, fix drift, reconcile removals, reset delta token). No parameters required.

```yaml
service: google_outlook_contacts_sync.full_sync
```

---

## Re-authorization

If either token expires or is revoked, the integration raises a **Repair** notification in HA. Go to **Settings → Devices & Services**, find the integration, and click **Re-configure** to re-authorize.

---

## Security

- **No secrets are stored in this repository.** All credentials are entered via the HA config flow and stored in HA's `.storage/` directory on your device.
- **`.storage/` is not encrypted.** Home Assistant writes config entries and this integration's mapping as plain JSON, protected only by file-system permissions. Your Microsoft and Google client secrets and OAuth refresh tokens are readable by anything that can read the config directory — **including any backup of it**. Treat backups of your HA configuration as secrets.
- Token caches, `.env` files, and `credentials.json` are listed in `.gitignore`.
- `detect-secrets` and Gitleaks run as pre-commit hooks.

---

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md).

---

## Disclaimer — trademarks and API compliance

This project is an unofficial, community-maintained integration and is **not
affiliated with, endorsed by, or sponsored by** Google LLC or Microsoft
Corporation. "Google", "Google Contacts", "Outlook" and "Microsoft" are
trademarks of their respective owners and are used here for descriptive
purposes only.

Each user creates and uses their **own** Google Cloud and Microsoft Entra OAuth
credentials. Responsibility for complying with the [Google API Services User
Data Policy](https://developers.google.com/terms/api-services-user-data-policy)
— including any OAuth app verification it requires — and with Microsoft's terms
lies with the user operating those credentials.

The integration icon is original artwork and is not derived from any
third-party logo.

---

## License

[MIT](LICENSE) © 2026 N1k4G
