"""DataUpdateCoordinator for periodic contact sync."""

from __future__ import annotations

import logging
from datetime import timedelta

from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.exceptions import ConfigEntryAuthFailed
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed

from .const import (
    CONF_SYNC_INTERVAL_HOURS,
    DATA_REAUTH_PROVIDER,
    DEFAULT_SYNC_INTERVAL_HOURS,
    DOMAIN,
)
from .sync.engine import SyncEngine, SyncPlan, SyncResult
from .sync.google_client import GoogleAuthError
from .sync.graph_client import MSAuthError

_LOGGER = logging.getLogger(__name__)


class ContactSyncCoordinator(DataUpdateCoordinator[SyncResult]):
    """Drives periodic sync via :class:`~.sync.engine.SyncEngine`."""

    def __init__(
        self,
        hass: HomeAssistant,
        entry: ConfigEntry,
        engine: SyncEngine,
    ) -> None:
        interval_hours: int = entry.options.get(
            CONF_SYNC_INTERVAL_HOURS,
            entry.data.get(CONF_SYNC_INTERVAL_HOURS, DEFAULT_SYNC_INTERVAL_HOURS),
        )
        super().__init__(
            hass,
            _LOGGER,
            name=f"google_outlook_contacts_sync ({entry.entry_id})",
            update_interval=timedelta(hours=interval_hours),
            config_entry=entry,
        )
        self._entry = entry
        self._engine = engine

    def _store_reauth_provider(self, exc: Exception) -> None:
        provider = "google" if isinstance(exc, GoogleAuthError) else "ms"
        self.hass.data.setdefault(DOMAIN, {}).setdefault(
            DATA_REAUTH_PROVIDER, {}
        )[self._entry.entry_id] = provider

    async def _async_update_data(self) -> SyncResult:
        """Fetch the latest data — called by the coordinator on schedule."""
        try:
            return await self._engine.async_sync()
        except (GoogleAuthError, MSAuthError) as exc:
            self._store_reauth_provider(exc)
            raise ConfigEntryAuthFailed(str(exc)) from exc
        except Exception as exc:
            raise UpdateFailed(f"Sync failed: {exc}") from exc

    async def async_full_resync(self) -> None:
        """Run a full reconciliation and publish the result to entities."""
        try:
            result = await self._engine.async_full_sync()
        except (GoogleAuthError, MSAuthError) as exc:
            self._store_reauth_provider(exc)
            self.async_set_update_error(UpdateFailed(str(exc)))
            if hasattr(self._entry, "async_start_reauth"):
                self._entry.async_start_reauth(self.hass)
            return
        except Exception as exc:
            self.async_set_update_error(UpdateFailed(f"Full resync failed: {exc}"))
            return
        self.async_set_updated_data(result)


class DryRunCoordinator(DataUpdateCoordinator[SyncPlan]):
    """Computes a dry-run preview on demand (no schedule).

    Refreshed manually by the dry-run button; ``update_interval`` is None so it
    never runs automatically. Applies nothing to Outlook.
    """

    def __init__(
        self,
        hass: HomeAssistant,
        entry: ConfigEntry,
        engine: SyncEngine,
    ) -> None:
        super().__init__(
            hass,
            _LOGGER,
            name=f"google_outlook_contacts_sync dry-run ({entry.entry_id})",
            update_interval=None,
        )
        self._engine = engine

    async def _async_update_data(self) -> SyncPlan:
        try:
            return await self._engine.async_plan()
        except Exception as exc:
            raise UpdateFailed(f"Dry run failed: {exc}") from exc
