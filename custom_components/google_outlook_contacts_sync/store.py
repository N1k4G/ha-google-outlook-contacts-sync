"""Persistent storage for the contact ID mapping and Google sync token."""

from __future__ import annotations

import logging
from typing import Any, cast

from homeassistant.core import HomeAssistant
from homeassistant.helpers.storage import Store

from .const import STORAGE_KEY_MAPPING, STORAGE_VERSION

_LOGGER = logging.getLogger(__name__)


def _new_default_data() -> dict[str, Any]:
    """Return a fresh default state.

    Built per call rather than copied from a module-level constant: a shallow
    ``dict()`` copy of a nested default leaves every instance sharing the same
    ``mapping``/``hashes`` objects, so one store's writes leak into stores
    created later in the same process.
    """
    return {
        "mapping": {},  # google_resource_name → graph_contact_id
        "hashes": {},  # google_resource_name → contact_hash
        "sync_token": None,  # Google People API syncToken
    }


class ContactSyncStore:
    """Persists the resourceName → Graph contact_id mapping and syncToken.

    Data is written to HA's ``.storage/`` directory as plain JSON, protected
    only by file-system permissions.

    The mapping is one-to-one in both directions: a Graph contact id must never
    be claimed by two different Google resource names, or both Google contacts
    would repeatedly overwrite the same Outlook contact.
    """

    def __init__(self, hass: HomeAssistant) -> None:
        self._store: Store[dict[str, Any]] = Store(
            hass, STORAGE_VERSION, STORAGE_KEY_MAPPING
        )
        self._data: dict[str, Any] = _new_default_data()

    async def async_load(self) -> None:
        """Load persisted data from storage (call once at setup)."""
        stored = await self._store.async_load()
        if not stored:
            return

        # Copy every nested container so the instance never aliases the object
        # handed over by the storage layer.
        data = _new_default_data()
        data["mapping"] = dict(stored.get("mapping") or {})
        data["hashes"] = dict(stored.get("hashes") or {})
        sync_token = stored.get("sync_token")
        data["sync_token"] = sync_token if isinstance(sync_token, str) else None
        self._data = data

        _LOGGER.debug(
            "Loaded %d contact mappings from storage", len(self._data["mapping"])
        )
        if conflicts := self.find_mapping_conflicts():
            _LOGGER.warning(
                "Storage contains %d Outlook contact(s) claimed by multiple Google "
                "contacts; these will not be updated until resolved: %s",
                len(conflicts),
                ", ".join(sorted(conflicts)),
            )

    async def async_save(self) -> None:
        """Persist current data to storage."""
        await self._store.async_save(self._data)

    async def async_remove(self) -> None:
        """Delete the backing storage file.

        Only called when the config entry is removed — a plain unload or reload
        must keep the mapping, otherwise every restart would re-adopt or
        re-create every contact.
        """
        await self._store.async_remove()

    # ------------------------------------------------------------------
    # Mapping operations
    # ------------------------------------------------------------------

    def lookup(self, resource_name: str) -> str | None:
        """Return the Graph contact id for a Google resource name, or None."""
        return cast("str | None", self._data["mapping"].get(resource_name))

    def save_mapping(
        self, resource_name: str, contact_id: str, contact_hash: str
    ) -> None:
        """Record a resource_name → contact_id mapping and its content hash.

        Raises :class:`MappingConflictError` when *contact_id* is already claimed
        by a different Google resource name. Re-saving an existing pair is always
        allowed, including when the store already holds a conflict from an
        earlier buggy run — refusing that would strand both contacts forever.
        """
        if self._data["mapping"].get(resource_name) != contact_id:
            conflicting = self.resource_names_for_contact_id(contact_id) - {
                resource_name
            }
            if conflicting:
                raise MappingConflictError(
                    f"Outlook contact {contact_id} is already mapped to "
                    f"{', '.join(sorted(conflicting))}; refusing to also map it to "
                    f"{resource_name}"
                )
        self._data["mapping"][resource_name] = contact_id
        self._data["hashes"][resource_name] = contact_hash

    def delete_mapping(self, resource_name: str) -> None:
        """Remove a mapping entry (contact was deleted or unlinked)."""
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

    def resource_names_for_contact_id(self, contact_id: str) -> set[str]:
        """Return every Google resource name mapped to *contact_id*."""
        return {
            resource_name
            for resource_name, mapped_id in self._data["mapping"].items()
            if mapped_id == contact_id
        }

    def find_mapping_conflicts(self) -> dict[str, set[str]]:
        """Return contact_id → resource names, for ids claimed more than once.

        Detects many-to-one corruption left behind by earlier versions; an empty
        dict means the one-to-one invariant holds.
        """
        by_contact_id: dict[str, set[str]] = {}
        for resource_name, contact_id in self._data["mapping"].items():
            by_contact_id.setdefault(contact_id, set()).add(resource_name)
        return {
            contact_id: resource_names
            for contact_id, resource_names in by_contact_id.items()
            if len(resource_names) > 1
        }

    def repoint_mapping(self, old_contact_id: str, new_contact_id: str) -> None:
        """Repoint every mapping entry from old_contact_id to new_contact_id.

        Used after duplicate resolution: when a duplicate Outlook contact is
        deleted, any Google resource that mapped to it is moved to the kept one.
        Callers must ensure this cannot merge two distinct Google resources onto
        one Outlook contact.
        """
        for resource_name, contact_id in self._data["mapping"].items():
            if contact_id == old_contact_id:
                self._data["mapping"][resource_name] = new_contact_id

    # ------------------------------------------------------------------
    # Sync token
    # ------------------------------------------------------------------

    @property
    def sync_token(self) -> str | None:
        return cast("str | None", self._data.get("sync_token"))

    @sync_token.setter
    def sync_token(self, value: str | None) -> None:
        self._data["sync_token"] = value


class MappingConflictError(Exception):
    """Raised when a Graph contact id would be claimed by two Google contacts."""
