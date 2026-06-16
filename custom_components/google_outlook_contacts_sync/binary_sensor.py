"""Binary sensors for the Google Outlook Contacts Sync integration."""

from __future__ import annotations

from typing import Any

from homeassistant.components.binary_sensor import (
    BinarySensorDeviceClass,
    BinarySensorEntity,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity import EntityCategory  # type: ignore[attr-defined]
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from .const import DATA_COORDINATOR, DOMAIN
from .coordinator import ContactSyncCoordinator
from .entity import build_device_info


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up binary sensors from a config entry."""
    data = hass.data[DOMAIN][entry.entry_id]
    coordinator: ContactSyncCoordinator = data[DATA_COORDINATOR]
    async_add_entities([SyncProblemBinarySensor(coordinator, entry)])


class SyncProblemBinarySensor(
    CoordinatorEntity[ContactSyncCoordinator], BinarySensorEntity
):
    """On when the last sync failed or reported per-contact failures."""

    _attr_has_entity_name = True
    _attr_translation_key = "sync_problem"
    _attr_device_class = BinarySensorDeviceClass.PROBLEM
    _attr_entity_category = EntityCategory.DIAGNOSTIC

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator)
        self._attr_unique_id = f"{entry.entry_id}_sync_problem"
        self._attr_device_info = build_device_info(entry)

    @property
    def is_on(self) -> bool:
        if not self.coordinator.last_update_success:
            return True
        data = self.coordinator.data
        return bool(data and data.failed > 0)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        attrs: dict[str, Any] = {}
        if self.coordinator.last_exception is not None:
            attrs["last_exception"] = str(self.coordinator.last_exception)
        data = self.coordinator.data
        if data and data.errors:
            attrs["errors"] = data.errors
        return attrs
