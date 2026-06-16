"""Diagnostics support for Google Outlook Contacts Sync.

All credentials and tokens are redacted — the output is safe to attach to a
GitHub issue.
"""

from __future__ import annotations

from typing import Any

from homeassistant.components.diagnostics import async_redact_data
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant

from .const import (
    CONF_GOOGLE_CLIENT_ID,
    CONF_GOOGLE_CLIENT_SECRET,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    DATA_COORDINATOR,
    DATA_STORE,
    DOMAIN,
)
from .coordinator import ContactSyncCoordinator
from .store import ContactSyncStore

_REDACT = {
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    CONF_GOOGLE_CLIENT_ID,
    CONF_GOOGLE_CLIENT_SECRET,
    "ms_token",
    "google_token",
}


async def async_get_config_entry_diagnostics(
    hass: HomeAssistant, entry: ConfigEntry
) -> dict[str, Any]:
    """Return diagnostics for a config entry."""
    data = hass.data[DOMAIN][entry.entry_id]
    coordinator: ContactSyncCoordinator = data[DATA_COORDINATOR]
    store: ContactSyncStore = data[DATA_STORE]

    last_result = coordinator.data.as_dict() if coordinator.data else None

    return {
        "entry": {
            "data": async_redact_data(dict(entry.data), _REDACT),
            "options": dict(entry.options),
        },
        "coordinator": {
            "last_update_success": coordinator.last_update_success,
            "last_exception": (
                str(coordinator.last_exception)
                if coordinator.last_exception
                else None
            ),
            "update_interval_seconds": (
                coordinator.update_interval.total_seconds()
                if coordinator.update_interval
                else None
            ),
            "last_result": last_result,
        },
        "store": {
            "synced_contacts": len(store.all_resource_names()),
            "has_sync_token": store.sync_token is not None,
        },
    }
