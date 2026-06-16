"""Shared entity helpers for the Google Outlook Contacts Sync integration."""

from __future__ import annotations

from homeassistant.config_entries import ConfigEntry
from homeassistant.helpers.device_registry import DeviceEntryType, DeviceInfo

from .const import DEVICE_MANUFACTURER, DEVICE_MODEL, DEVICE_NAME, DOMAIN


def build_device_info(entry: ConfigEntry) -> DeviceInfo:
    """Return the DeviceInfo that groups all entities under one HA device."""
    return DeviceInfo(
        identifiers={(DOMAIN, entry.entry_id)},
        name=DEVICE_NAME,
        manufacturer=DEVICE_MANUFACTURER,
        model=DEVICE_MODEL,
        entry_type=DeviceEntryType.SERVICE,
    )
