"""Persistent storage for the contact ID mapping and Google sync token."""

from __future__ import annotations

import logging
from typing import Any, cast

from homeassistant.core import HomeAssistant
from homeassistant.helpers.storage import Store

from .const import STORAGE_KEY_MAPPING, STORAGE_VERSION

_LOGGER = logging.getLogger(__name__)

_DEFAULT_DATA: dict[str, Any] = {
    "mapping": {},       # google_resource_name → graph_contact_id
    "hashes": {},        # google_resource_name → contact_hash
    "sync_token": None,  # Google People API syncToken
}


class ContactSyncStore:
    """Persists the resourceName → Graph contact_id mapping and syncToken.

    Data is written to HA's encrypted ``.storage/`` directory.
    """

    def __init__(self, hass: HomeAssistant) -> None:
        self._store: Store[dict[str, Any]] = Store(
            hass, STORAGE_VERSION, STORAGE_KEY_MAPPING
        )
        self._data: dict[str, Any] = dict(_DEFAULT_DATA)

    async def async_load(self) -> None:
        """Load persisted data from storage (call once at setup)."""
        stored = await self._store.async_load()
        if stored:
            self._data = {**_DEFAULT_DATA, **stored}
            _LOGGER.debug(
                "Loaded %d contact mappings from storage",
                len(self._data["mapping"]),
            )

    async def async_save(self) -> None:
        """Persist current data to storage."""
        await self._store.async_save(self._data)

    # ------------------------------------------------------------------
    # Mapping operations
    # ------------------------------------------------------------------

    def lookup(self, resource_name: str) -> str | None:
        """Return the Graph contact id for a Google resource name, or None."""
        return cast("str | None", self._data["mapping"].get(resource_name))

    def save_mapping(
        self, resource_name: str, contact_id: str, contact_hash: str
    ) -> None:
        """Record a resource_name → contact_id mapping and its content hash."""
        self._data["mapping"][resource_name] = contact_id
        self._data["hashes"][resource_name] = contact_hash

    def delete_mapping(self, resource_name: str) -> None:
        """Remove a mapping entry (contact was deleted)."""
        self._data["mapping"].pop(resource_name, None)
        self._data["hashes"].pop(resource_name, None)

    def get_hash(self, resource_name: str) -> str | None:
        """Return the last-seen content hash for a contact, or None."""
        return cast("str | None", self._data["hashes"].get(resource_name))

    def all_resource_names(self) -> set[str]:
        """Return all tracked Google resource names."""
        return set(self._data["mapping"].keys())

    def mapped_ids(self) -> set[str]:
        """Return all Graph contact ids currently referenced by the mapping."""
        return set(self._data["mapping"].values())

    def repoint_mapping(self, old_contact_id: str, new_contact_id: str) -> None:
        """Repoint every mapping entry from old_contact_id to new_contact_id.

        Used after duplicate resolution: when a duplicate Outlook contact is
        deleted, any Google resource that mapped to it is moved to the kept one.
        """
        for resource_name, contact_id in self._data["mapping"].items():
            if contact_id == old_contact_id:
                self._data["mapping"][resource_name] = new_contact_id

    # ------------------------------------------------------------------
    # Sync token
    # ------------------------------------------------------------------

    @property
    def sync_token(self) -> str | None:
        return self._data.get("sync_token")

    @sync_token.setter
    def sync_token(self, value: str | None) -> None:
        self._data["sync_token"] = value
