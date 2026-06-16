"""Buttons for the Google Outlook Contacts Sync integration."""

from __future__ import annotations

from homeassistant.components.button import ButtonEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback

from .const import DATA_COORDINATOR, DATA_DRY_RUN_COORDINATOR, DOMAIN
from .coordinator import ContactSyncCoordinator, DryRunCoordinator
from .entity import build_device_info


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up buttons from a config entry."""
    data = hass.data[DOMAIN][entry.entry_id]
    coordinator: ContactSyncCoordinator = data[DATA_COORDINATOR]
    dry_run_coordinator: DryRunCoordinator = data[DATA_DRY_RUN_COORDINATOR]

    async_add_entities(
        [
            SyncNowButton(coordinator, entry),
            FullResyncButton(coordinator, entry),
            DryRunButton(dry_run_coordinator, entry),
        ]
    )


class SyncNowButton(ButtonEntity):
    """Trigger an immediate delta sync."""

    _attr_has_entity_name = True
    _attr_translation_key = "sync_now"

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        self._coordinator = coordinator
        self._attr_unique_id = f"{entry.entry_id}_sync_now"
        self._attr_device_info = build_device_info(entry)

    async def async_press(self) -> None:
        await self._coordinator.async_request_refresh()


class FullResyncButton(ButtonEntity):
    """Run a full reconciliation (adopt existing Outlook contacts, fix drift)."""

    _attr_has_entity_name = True
    _attr_translation_key = "full_resync"

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        self._coordinator = coordinator
        self._attr_unique_id = f"{entry.entry_id}_full_resync"
        self._attr_device_info = build_device_info(entry)

    async def async_press(self) -> None:
        await self._coordinator.async_full_resync()


class DryRunButton(ButtonEntity):
    """Compute a dry-run preview without changing Outlook."""

    _attr_has_entity_name = True
    _attr_translation_key = "run_dry_run"

    def __init__(
        self, coordinator: DryRunCoordinator, entry: ConfigEntry
    ) -> None:
        self._coordinator = coordinator
        self._attr_unique_id = f"{entry.entry_id}_run_dry_run"
        self._attr_device_info = build_device_info(entry)

    async def async_press(self) -> None:
        await self._coordinator.async_request_refresh()
