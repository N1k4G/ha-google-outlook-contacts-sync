# Microsoft device-code rollout

Implements the software portion of [issue #41](https://github.com/N1k4G/ha-google-outlook-contacts-sync/issues/41).
The Entra registration setting and HA integration release are separate
rollout steps. Keep the application secret and previous callback until the
checklist below is complete for every consumer.

## Order

1. Review/apply the additive infrastructure PR enabling
   `fallback_public_client_enabled` on the existing application. The app ID,
   delegated Graph scopes, service principal, web callback and password remain.
2. Install the integration release and restart HA. Existing entries default to
   their previous confidential mode; no automatic login migration occurs.
3. Use **Reconfigure**, retain the IDs and Google settings, choose **Device
   code**, and complete a fresh Microsoft sign-in. Do not reuse a confidential
   refresh token as proof that a public login works.
4. HA polls with bounded HTTP timeouts, honors pending/slow-down intervals, and
   verifies a forced refresh before committing the new cache and mode. Closing
   the flow, shutdown, entry unload, timeout or failed login saves no changes.
5. After successful reconfiguration, verify the live acceptance checklist.
6. Only after all consumers migrate, make a separate cleanup PR to remove the
   Entra application password and sensitive Terraform outputs. Historical state
   and backup versions may still contain old credentials.

## Live acceptance (requires the real personal Microsoft account)

- [ ] HCP plan preserves the application and password; only the public-client
      setting changes.
- [ ] First HA install resolves the pinned O365/MSAL requirements without
      conflicts with other installed integrations.
- [ ] Device login succeeds with the existing app ID and delegated permissions.
- [ ] HA reconfiguration succeeds and the entry no longer has a Microsoft
      secret. Do not paste entry data into logs or an issue.
- [ ] First sync and forced token refresh succeed.
- [ ] Restart HA; verify scheduled sync and a controlled contact update.
- [ ] Google login, options, mapping and entity identities are retained.
- [ ] Failed/cancelled login leaves the working entry usable.
- [ ] All other consumers of the old application secret have migrated.
- [ ] Secret cleanup is reviewed separately after these checks.

For legacy caches without an `Account` identity, the integration cannot
compare the old and new Microsoft account automatically. Explicitly verify
that the same personal Microsoft account is used during the live migration;
existing contact mapping IDs belong to that original Outlook account.

A deployment agent without account access must leave the live boxes open and
retain the old Entra secret. Offline SDK and flow tests do not establish real
consent, Entra registration acceptance, or account-specific refresh policy.

## Dependency and test contract

O365 2.1.9 and MSAL 1.37.0 are pinned because this implementation depends on
single-attempt `acquire_token_by_device_flow(..., exit_condition=...)`, the MSAL
cache layout and O365 refresh behavior. Update both with the real-SDK contract
tests. The device flow and cache are ephemeral; codes are not stored in entries
or diagnostics. The cache backend strips one-time protocol secrets from SDK
cache-event debug data while retaining fields needed for caching.

Run the complete tests and coverage gate in Linux CI, plus Ruff and mypy. The
synchronous/mocked boundaries can also run on Windows with only pytest-asyncio
and pytest-cov loaded when HA's full test plugin requires Linux-only `fcntl`.
