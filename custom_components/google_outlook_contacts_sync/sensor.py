"""Sensors for the Google Outlook Contacts Sync integration."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from homeassistant.components.sensor import (
    SensorDeviceClass,
    SensorEntity,
    SensorStateClass,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity import EntityCategory  # type: ignore[attr-defined]
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from .const import (
    DATA_COORDINATOR,
    DATA_DRY_RUN_COORDINATOR,
    DATA_STORE,
    DOMAIN,
)
from .coordinator import ContactSyncCoordinator, DryRunCoordinator
from .entity import build_device_info
from .store import ContactSyncStore

# Cap the number of names exposed per dry-run bucket to keep attributes sane.
_MAX_PREVIEW_ITEMS = 200


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up sensors from a config entry."""
    data = hass.data[DOMAIN][entry.entry_id]
    coordinator: ContactSyncCoordinator = data[DATA_COORDINATOR]
    store: ContactSyncStore = data[DATA_STORE]

    dry_run_coordinator: DryRunCoordinator = data[DATA_DRY_RUN_COORDINATOR]

    async_add_entities(
        [
            LastSyncSensor(coordinator, entry),
            LastSyncResultSensor(coordinator, entry),
            NextSyncSensor(coordinator, entry),
            SyncedContactsSensor(coordinator, entry, store),
            DuplicatesRemovedSensor(coordinator, entry),
            DryRunPreviewSensor(dry_run_coordinator, entry),
        ]
    )


class _BaseSyncSensor(CoordinatorEntity[ContactSyncCoordinator], SensorEntity):
    """Common wiring for sensors backed by the main sync coordinator."""

    _attr_has_entity_name = True

    def __init__(
        self,
        coordinator: ContactSyncCoordinator,
        entry: ConfigEntry,
        key: str,
    ) -> None:
        super().__init__(coordinator)
        self._attr_translation_key = key
        self._attr_unique_id = f"{entry.entry_id}_{key}"
        self._attr_device_info = build_device_info(entry)


class LastSyncSensor(_BaseSyncSensor):
    """Timestamp of the last successful sync."""

    _attr_device_class = SensorDeviceClass.TIMESTAMP

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator, entry, "last_sync")

    @property
    def native_value(self) -> datetime | None:
        if self.coordinator.data is None:
            return None
        return self.coordinator.data.finished_at


class LastSyncResultSensor(_BaseSyncSensor):
    """Verbose summary of the most recent sync run."""

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator, entry, "last_sync_result")

    @property
    def native_value(self) -> str | None:
        if self.coordinator.data is None:
            return None
        return self.coordinator.data.summary()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        result = self.coordinator.data
        if result is None:
            return {}
        return {
            "created": result.created,
            "updated": result.updated,
            "deleted": result.deleted,
            "skipped": result.skipped,
            "failed": result.failed,
            "duration_seconds": result.duration_seconds,
            "errors": result.errors,
        }


class NextSyncSensor(_BaseSyncSensor):
    """Estimated timestamp of the next scheduled sync."""

    _attr_device_class = SensorDeviceClass.TIMESTAMP
    _attr_entity_category = EntityCategory.DIAGNOSTIC

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator, entry, "next_sync")

    @property
    def native_value(self) -> datetime | None:
        data = self.coordinator.data
        interval = self.coordinator.update_interval
        if data is None or data.finished_at is None or interval is None:
            return None
        return data.finished_at + interval


class SyncedContactsSensor(_BaseSyncSensor):
    """Number of contacts currently tracked/synced to Outlook."""

    _attr_state_class = SensorStateClass.MEASUREMENT
    _attr_native_unit_of_measurement = "contacts"

    def __init__(
        self,
        coordinator: ContactSyncCoordinator,
        entry: ConfigEntry,
        store: ContactSyncStore,
    ) -> None:
        super().__init__(coordinator, entry, "synced_contacts")
        self._store = store

    @property
    def native_value(self) -> int:
        return len(self._store.all_resource_names())


class DuplicatesRemovedSensor(_BaseSyncSensor):
    """Number of duplicate Outlook contacts removed in the last sync."""

    _attr_state_class = SensorStateClass.MEASUREMENT
    _attr_native_unit_of_measurement = "contacts"
    _attr_entity_category = EntityCategory.DIAGNOSTIC

    def __init__(
        self, coordinator: ContactSyncCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator, entry, "duplicates_removed")

    @property
    def native_value(self) -> int | None:
        if self.coordinator.data is None:
            return None
        return self.coordinator.data.duplicates_removed

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        data = self.coordinator.data
        if data is None or not data.duplicates:
            return {}
        return {"groups": data.duplicates}


class DryRunPreviewSensor(
    CoordinatorEntity[DryRunCoordinator], SensorEntity
):
    """Preview of the contacts the next sync would add/update/delete."""

    _attr_has_entity_name = True
    _attr_native_unit_of_measurement = "changes"
    _attr_state_class = SensorStateClass.MEASUREMENT

    def __init__(
        self, coordinator: DryRunCoordinator, entry: ConfigEntry
    ) -> None:
        super().__init__(coordinator)
        self._attr_translation_key = "dry_run_preview"
        self._attr_unique_id = f"{entry.entry_id}_dry_run_preview"
        self._attr_device_info = build_device_info(entry)

    @property
    def native_value(self) -> int | None:
        if self.coordinator.data is None:
            return None
        return self.coordinator.data.total

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        plan = self.coordinator.data
        if plan is None:
            return {}
        return {
            "to_create": [c.display_name for c in plan.to_create[:_MAX_PREVIEW_ITEMS]],
            "to_update": [c.display_name for c in plan.to_update[:_MAX_PREVIEW_ITEMS]],
            "to_delete": [c.display_name for c in plan.to_delete[:_MAX_PREVIEW_ITEMS]],
            "create_count": len(plan.to_create),
            "update_count": len(plan.to_update),
            "delete_count": len(plan.to_delete),
            "computed_at": (
                plan.finished_at.isoformat() if plan.finished_at else None
            ),
        }
