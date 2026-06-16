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

from .auth.ms_auth import build_account
from .const import (
    CONF_AUTO_REMOVE_DUPLICATES,
    CONF_DELETE_REMOVED,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    DATA_COORDINATOR,
    DATA_DRY_RUN_COORDINATOR,
    DATA_STORE,
    DEFAULT_AUTO_REMOVE_DUPLICATES,
    DEFAULT_DELETE_REMOVED,
    DOMAIN,
    SERVICE_FULL_SYNC,
    SERVICE_SYNC_NOW,
)
from .coordinator import ContactSyncCoordinator, DryRunCoordinator
from .http import MSAuthCallbackView
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

# Required by hassfest when async_setup is present with no configuration parameters.
CONFIG_SCHEMA = cv.config_entry_only_config_schema(DOMAIN)


async def async_setup(hass: HomeAssistant, config: ConfigType) -> bool:
    """Register the MS OAuth callback view at HA startup."""
    hass.http.register_view(MSAuthCallbackView(hass))
    hass.data.setdefault(DOMAIN, {})
    return True


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up the integration from a config entry."""
    data = entry.data

    google_token: dict[str, Any] = data.get("google_token", {})
    google_client = await hass.async_add_executor_job(
        GoogleContactsClient.from_token_dict, google_token
    )

    ms_token: dict[str, Any] = data.get("ms_token", {})
    account, ms_backend = await hass.async_add_executor_job(
        build_account,
        data[CONF_MS_CLIENT_ID],
        data[CONF_MS_CLIENT_SECRET],
        ms_token,
    )

    if not await hass.async_add_executor_job(lambda: account.is_authenticated):
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
    engine = SyncEngine(
        hass,
        google_client,
        graph_client,
        store,
        delete_removed,
        auto_remove_duplicates,
    )
    coordinator = ContactSyncCoordinator(hass, entry, engine)
    dry_run_coordinator = DryRunCoordinator(hass, entry, engine)

    await coordinator.async_config_entry_first_refresh()

    hass.data[DOMAIN][entry.entry_id] = {
        DATA_COORDINATOR: coordinator,
        DATA_DRY_RUN_COORDINATOR: dry_run_coordinator,
        DATA_STORE: store,
    }

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)

    async def _handle_sync_now(_call: ServiceCall) -> None:
        await coordinator.async_refresh()

    async def _handle_full_sync(_call: ServiceCall) -> None:
        await coordinator.async_full_resync()

    hass.services.async_register(DOMAIN, SERVICE_SYNC_NOW, _handle_sync_now)
    hass.services.async_register(DOMAIN, SERVICE_FULL_SYNC, _handle_full_sync)

    entry.async_on_unload(entry.add_update_listener(_async_options_updated))

    return True


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Unload a config entry."""
    unloaded = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    if unloaded:
        hass.data[DOMAIN].pop(entry.entry_id, None)
    return unloaded


async def _async_options_updated(hass: HomeAssistant, entry: ConfigEntry) -> None:
    """Reload the entry when options change so the new interval takes effect."""
    await hass.config_entries.async_reload(entry.entry_id)
