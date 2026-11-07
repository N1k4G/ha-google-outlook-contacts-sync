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

1. In [Entra App registrations](https://portal.azure.com), create an application
   supporting **Accounts in any organizational directory and personal Microsoft
   accounts**. Keep the Application (client) ID.
2. Add Microsoft Graph **delegated** permissions `Contacts.ReadWrite` and
   `User.Read`. The integration signs in the user; it does not use app-only
   permissions.
3. For the recommended **Device code** method, open **Authentication**, enable
   **Allow public client flows**, and save. No Microsoft client secret or
   Microsoft redirect URI is needed for this method.
4. For the optional previous **Login with client secret** method, register a
   **Web** redirect URI
   `https://<your-ha-external-url>/api/google_outlook_contacts_sync/auth` and
   create a client secret under **Certificates & secrets**.

If this registration is managed by
[N1k4G/infrastructure](https://github.com/N1k4G/infrastructure), apply its additive
`fallback_public_client_enabled = true` change through HCP before migrating HA.
Existing installations can keep using their previous login during the rollout.

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

**Requires Home Assistant 2025.8.0 or newer.** The integration uses
`OptionsFlowWithReload` (2025.8), `ConfigFlowResult` (2024.4), the
`config_entry` argument to `DataUpdateCoordinator` (2024.11/12) and the
`single_config_entry` manifest flag (2024.11); on older versions setup fails at
import.

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

1. In **Settings ? Devices & Services ? Add Integration**, search for
   "Google Outlook Contacts Sync".
2. Select **Device code (no secret, recommended)** and enter the Microsoft
   client ID. Leave the Microsoft secret empty.
3. Open the Microsoft sign-in link displayed by Home Assistant, enter the
   displayed code, and approve contacts access with your personal Microsoft
   account. HA waits for confirmation and verifies token refresh before saving.
   Only authorize a code you requested from your own HA instance.
4. Enter the Google Web OAuth client ID and secret, then complete Google
   sign-in in the browser. Google's existing external URL and callback
   requirements still apply.
5. The first full sync runs immediately; subsequent runs use delta sync.

The previous Microsoft secret-based method is still available. It uses the
existing browser callback instead of a device code.

### Migrate an existing Microsoft login

Open the integration menu and choose **Reconfigure**. Keep the client IDs,
select **Device code**, and submit. Complete a fresh Microsoft login with the
new code, using the same Microsoft account as before. Until login and its refresh check succeed, the stored credentials and
running configuration are unchanged. A failed or cancelled attempt saves
nothing. On success HA preserves Google authorization, the contact mapping,
options and entity identity, removes the Microsoft secret from the entry and
reloads the integration.

There is no annual Microsoft client-secret rotation in device-code mode.
Microsoft can still require a new sign-in after a revocation or other account
policy change; use HA's reauthentication flow when requested. Restarting HA
while a sign-in is pending discards that attempt; request a new code.

Keep the Entra application secret during rollout if other installations still
use it. Revoke it and remove Terraform secret resources/outputs only after all
consumers have migrated and live login, forced refresh, restart and scheduled
sync have passed. Removing the key from the current entry does not erase old
HA backups or Terraform state history. See the
[migration and live acceptance checklist](docs/microsoft-device-code-migration.md).

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

If either token expires or is revoked, the integration raises a **Repair** notification in HA. Go to **Settings → Devices & Services**, find the integration, and click **Reconfigure** on the re-authentication prompt. The prompt names the affected provider and accepts a **new client secret**: if the stored secret expired or was revoked, enter the replacement there, otherwise leave it empty to keep the stored one. You then sign in with that provider again.

---

## Rotating or changing client secrets

Client secrets and IDs can be changed at any time without removing the integration. Tokens, the contact mapping and the Google sync token are kept, so no full resync and no duplicates follow.

1. Create the new secret **while the old one still works**:
   - **Microsoft:** Entra ID → App registrations → your app → **Certificates & secrets → New client secret**. Copy the value.
   - **Google:** APIs & Services → Credentials → your OAuth client → **Add secret**. Copy the value.
2. In HA go to **Settings → Devices & Services → Google Outlook Contacts Sync → ⋮ → Reconfigure**.
3. Enter the new secret(s). Leave a secret field empty to keep the stored one. Before saving, the integration redeems your existing refresh token with the new secret:
   - **Accepted:** the secret is saved and the integration reloads. You do not need to sign in again.
   - **Rejected by the provider** (wrong or expired secret): the form shows an error and nothing is saved.
   - **Refresh token no longer usable:** you are taken to that provider's sign-in, as in a re-authorization.
4. Press **Sync now** and check that **Sync problem** stays off.
5. Delete the old secret in Entra or Google Cloud.

Changing a **client ID** (moving to a different app registration or OAuth client) works the same way, but it always requires that client's secret and a fresh sign-in with the provider, because refresh tokens are bound to the client that issued them. Sign in with the same Microsoft and Google accounts as before; with a different account the stored mapping no longer matches.

If you suspect a secret has leaked, for example through a backup, rotate it this way right away and then delete the old secret at the provider.

---

## Security

- **No secrets are stored in this repository.** All credentials are entered via the HA config flow and stored in HA's `.storage/` directory on your device.
- **`.storage/` is not encrypted.** Home Assistant writes config entries and this integration's mapping as plain JSON, protected only by file-system permissions. Your Microsoft and Google client secrets and OAuth refresh tokens are readable by anything that can read the config directory — **including any backup of it**. Treat backups of your HA configuration as secrets, and [rotate the client secrets](#rotating-or-changing-client-secrets) if one may have been exposed.
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
