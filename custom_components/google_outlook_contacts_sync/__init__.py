"""Google Outlook Contacts Sync — Home Assistant custom integration."""

from __future__ import annotations

import logging
from typing import Any

import homeassistant.helpers.config_validation as cv
from homeassistant.config_entries import ConfigEntry
from homeassistant.const import Platform
from homeassistant.core import HomeAssistant, ServiceCall
from homeassistant.exceptions import ConfigEntryAuthFailed
from homeassistant.helpers.typing import ConfigType

from .auth.ms_auth import HAStoreTokenBackend, build_account
from .const import (
    CONF_AUTO_REMOVE_DUPLICATES,
    CONF_DELETE_REMOVED,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    DATA_COORDINATOR,
    DATA_DRY_RUN_COORDINATOR,
    DATA_OPTIONS_SNAPSHOT,
    DATA_REAUTH_PROVIDER,
    DATA_STORE,
    DEFAULT_AUTO_REMOVE_DUPLICATES,
    DEFAULT_DELETE_REMOVED,
    DOMAIN,
    SERVICE_FULL_SYNC,
    SERVICE_SYNC_NOW,
)
from .coordinator import ContactSyncCoordinator, DryRunCoordinator
from .http import async_register_callback_views
from .store import ContactSyncStore
from .sync.engine import SyncEngine
from .sync.google_client import GoogleContactsClient
from .sync.graph_client import GraphContactsClient

_LOGGER = logging.getLogger(__name__)

PLATFORMS: list[Platform] = [
    Platform.BINARY_SENSOR,
    Platform.BUTTON,
    Platform.SENSOR,
]

# Everything needed to rebuild a Google client; without these the stored token
# is unusable and only a fresh authorization can help.
_REQUIRED_GOOGLE_TOKEN_KEYS = ("client_id", "client_secret", "refresh_token")

# Required by hassfest when async_setup is present with no configuration parameters.
CONFIG_SCHEMA = cv.config_entry_only_config_schema(DOMAIN)


async def async_setup(hass: HomeAssistant, config: ConfigType) -> bool:
    """Register the OAuth callback views at HA startup."""
    hass.data.setdefault(DOMAIN, {})
    async_register_callback_views(hass)
    return True


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up the integration from a config entry.

    Deliberately does not wait for a sync. Entities are published first and the
    initial run happens in the background, so a large address book or a slow API
    cannot stall startup and a transient sync error cannot fail the setup.
    """
    data = entry.data

    google_token: dict[str, Any] = data.get("google_token") or {}
    if not all(google_token.get(key) for key in _REQUIRED_GOOGLE_TOKEN_KEYS):
        _store_reauth_provider(hass, entry, "google")
        raise ConfigEntryAuthFailed(
            "Google credentials are missing or incomplete. Please re-authorize."
        )

    google_client = await hass.async_add_executor_job(
        GoogleContactsClient.from_token_dict, google_token
    )

    ms_token: dict[str, Any] = data.get("ms_token") or {}
    account, ms_backend = await hass.async_add_executor_job(
        build_account,
        data[CONF_MS_CLIENT_ID],
        data[CONF_MS_CLIENT_SECRET],
        ms_token,
    )

    if not await hass.async_add_executor_job(lambda: account.is_authenticated):
        _store_reauth_provider(hass, entry, "ms")
        raise ConfigEntryAuthFailed(
            "Microsoft account is not authenticated. Please re-authorize."
        )

    graph_client = GraphContactsClient(account)

    store = ContactSyncStore(hass)
    await store.async_load()

    delete_removed: bool = entry.options.get(
        CONF_DELETE_REMOVED,
        entry.data.get(CONF_DELETE_REMOVED, DEFAULT_DELETE_REMOVED),
    )
    auto_remove_duplicates: bool = entry.options.get(
        CONF_AUTO_REMOVE_DUPLICATES, DEFAULT_AUTO_REMOVE_DUPLICATES
    )

    async def _persist_tokens() -> None:
        await _async_persist_tokens(hass, entry, ms_backend, google_client)

    engine = SyncEngine(
        hass,
        google_client,
        graph_client,
        store,
        delete_removed,
        auto_remove_duplicates,
        persist_tokens=_persist_tokens,
    )
    coordinator = ContactSyncCoordinator(hass, entry, engine)
    dry_run_coordinator = DryRunCoordinator(hass, entry, engine)

    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = {
        DATA_COORDINATOR: coordinator,
        DATA_DRY_RUN_COORDINATOR: dry_run_coordinator,
        DATA_STORE: store,
        DATA_OPTIONS_SNAPSHOT: dict(entry.options),
    }

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)

    entry.async_on_unload(entry.add_update_listener(_async_entry_updated))
    _async_register_services(hass)

    # No cursor means there is nothing to apply a delta against, so the first
    # run must reconcile: adopt matching Outlook contacts rather than duplicate
    # them. A restart with a cursor takes the cheap path.
    if store.sync_token is None:
        initial_sync = coordinator.async_full_resync()
    else:
        initial_sync = coordinator.async_request_refresh()

    entry.async_create_background_task(
        hass,
        initial_sync,
        f"{DOMAIN} initial sync",
    )

    return True


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Unload a config entry, leaving persisted data intact."""
    unloaded = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    if unloaded:
        hass.data[DOMAIN].pop(entry.entry_id, None)
        if not _entry_runtimes(hass):
            # The service closures would otherwise survive with no entry behind
            # them, still listed in the UI and pointing at a dead coordinator.
            hass.services.async_remove(DOMAIN, SERVICE_SYNC_NOW)
            hass.services.async_remove(DOMAIN, SERVICE_FULL_SYNC)
    return unloaded


async def async_remove_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Delete the persisted mapping when the integration is removed.

    Only on removal — async_unload_entry must leave the data alone, or every
    reload and restart would re-adopt or recreate every contact.
    """
    store = ContactSyncStore(hass)
    await store.async_remove()


async def _async_entry_updated(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Reload only when the user's options actually changed.

    HA calls update listeners for any entry change, including the token writes
    made after each run. Reloading on those would start a sync, which can
    refresh a token, which writes the entry again — a loop. Comparing against
    the snapshot taken at setup distinguishes the two deterministically, unlike
    an "ignore the next update" flag, which races when updates arrive quickly.
    """
    runtime = hass.data.get(DOMAIN, {}).get(entry.entry_id)
    if runtime is None:
        return
    if runtime.get(DATA_OPTIONS_SNAPSHOT) == dict(entry.options):
        return
    await hass.config_entries.async_reload(entry.entry_id)


async def _async_persist_tokens(
    hass: HomeAssistant,
    entry: ConfigEntry,
    ms_backend: HAStoreTokenBackend,
    google_client: GoogleContactsClient,
) -> None:
    """Write back OAuth state that was refreshed at runtime.

    Both clients renew tokens in memory only. Rebuilt from entry.data on the
    next restart, HA would load a stale cache and drop into reauth for
    credentials that are actually fine.

    Both providers are written in a single update so one run produces at most
    one entry write, and nothing is written when nothing changed.
    """
    new_data = dict(entry.data)
    changed = False

    current_ms = ms_backend.get_token_dict()
    if current_ms and current_ms != entry.data.get("ms_token"):
        new_data["ms_token"] = current_ms
        changed = True

    current_google = google_client.get_token_dict()
    if current_google and current_google != entry.data.get("google_token"):
        new_data["google_token"] = current_google
        changed = True

    if changed:
        _LOGGER.debug("Persisting refreshed OAuth token(s) to the config entry")
        hass.config_entries.async_update_entry(entry, data=new_data)


def _store_reauth_provider(
    hass: HomeAssistant, entry: ConfigEntry, provider: str
) -> None:
    hass.data.setdefault(DOMAIN, {}).setdefault(DATA_REAUTH_PROVIDER, {})[
        entry.entry_id
    ] = provider


def _entry_runtimes(hass: HomeAssistant) -> list[dict[str, Any]]:
    """Runtime data of every currently loaded entry.

    hass.data[DOMAIN] also holds non-entry keys (pending OAuth flows, the reauth
    provider map, the view guard), so entries are identified by shape.
    """
    return [
        value
        for key, value in hass.data.get(DOMAIN, {}).items()
        if isinstance(value, dict) and DATA_COORDINATOR in value
    ]


def _async_register_services(hass: HomeAssistant) -> None:
    """Register the domain services once, resolving coordinators at call time.

    Closing over one entry's coordinator would leave the services pointing at a
    stale object after a reload.
    """
    if hass.services.has_service(DOMAIN, SERVICE_SYNC_NOW):
        return

    async def _handle_sync_now(_call: ServiceCall) -> None:
        for runtime in _entry_runtimes(hass):
            # Same debounced path as the Sync now button.
            await runtime[DATA_COORDINATOR].async_request_refresh()

    async def _handle_full_sync(_call: ServiceCall) -> None:
        for runtime in _entry_runtimes(hass):
            await runtime[DATA_COORDINATOR].async_full_resync()

    hass.services.async_register(DOMAIN, SERVICE_SYNC_NOW, _handle_sync_now)
    hass.services.async_register(DOMAIN, SERVICE_FULL_SYNC, _handle_full_sync)
