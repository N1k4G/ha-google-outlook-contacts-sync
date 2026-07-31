"""Tests for store.py — instance isolation and the one-to-one mapping invariant."""

from __future__ import annotations

from typing import Any

import pytest

from custom_components.google_outlook_contacts_sync import store as store_module
from custom_components.google_outlook_contacts_sync.store import (
    ContactSyncStore,
    MappingConflictError,
)


class FakeHAStore:
    """Stands in for homeassistant.helpers.storage.Store."""

    def __init__(self, data: dict[str, Any] | None = None):
        self.data = data
        self.saved: dict[str, Any] | None = None
        self.removed = False

    async def async_load(self):
        return self.data

    async def async_save(self, data):
        self.saved = data

    async def async_remove(self):
        self.removed = True


def make_store(monkeypatch, stored: dict[str, Any] | None = None):
    """Build a ContactSyncStore backed by a FakeHAStore."""
    backing = FakeHAStore(stored)
    monkeypatch.setattr(store_module, "Store", lambda *a, **k: backing)
    return ContactSyncStore(hass=object()), backing  # type: ignore[arg-type]


class TestInstanceIsolation:
    async def test_two_fresh_stores_are_independent(self, monkeypatch):
        a, _ = make_store(monkeypatch)
        b, _ = make_store(monkeypatch)
        await a.async_load()
        await b.async_load()

        a.save_mapping("people/a", "graph-a", "hash-a")

        assert a.lookup("people/a") == "graph-a"
        assert b.lookup("people/a") is None
        assert b.all_resource_names() == set()

    async def test_store_created_later_inherits_nothing(self, monkeypatch):
        """Regression for #28: shallow-copied nested defaults leaked between stores."""
        a, _ = make_store(monkeypatch)
        await a.async_load()
        a.save_mapping("people/a", "graph-a", "hash-a")

        b, _ = make_store(monkeypatch)
        await b.async_load()

        assert b.all_resource_names() == set()
        assert b.mapped_ids() == set()

    async def test_hashes_are_independent(self, monkeypatch):
        a, _ = make_store(monkeypatch)
        b, _ = make_store(monkeypatch)
        await a.async_load()
        await b.async_load()

        a.save_mapping("people/a", "graph-a", "hash-a")

        assert a.get_hash("people/a") == "hash-a"
        assert b.get_hash("people/a") is None

    async def test_loaded_data_is_not_aliased_to_the_source_dict(self, monkeypatch):
        stored = {"mapping": {"people/a": "graph-a"}, "hashes": {"people/a": "h"}}
        a, _ = make_store(monkeypatch, stored)
        await a.async_load()

        a.save_mapping("people/b", "graph-b", "hash-b")

        assert stored["mapping"] == {"people/a": "graph-a"}
        assert stored["hashes"] == {"people/a": "h"}


class TestLoad:
    async def test_load_without_stored_data_keeps_empty_defaults(self, monkeypatch):
        s, _ = make_store(monkeypatch, None)
        await s.async_load()
        assert s.all_resource_names() == set()
        assert s.sync_token is None

    async def test_load_roundtrip(self, monkeypatch):
        stored = {
            "mapping": {"people/a": "graph-a"},
            "hashes": {"people/a": "hash-a"},
            "sync_token": "token-1",
        }
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        assert s.lookup("people/a") == "graph-a"
        assert s.get_hash("people/a") == "hash-a"
        assert s.sync_token == "token-1"

    async def test_load_normalizes_missing_and_invalid_fields(self, monkeypatch):
        s, _ = make_store(monkeypatch, {"sync_token": 12345})
        await s.async_load()

        assert s.all_resource_names() == set()
        assert s.get_hash("anything") is None
        assert s.sync_token is None

    async def test_save_writes_instance_local_state(self, monkeypatch):
        s, backing = make_store(monkeypatch)
        await s.async_load()
        s.save_mapping("people/a", "graph-a", "hash-a")
        s.sync_token = "token-2"
        await s.async_save()

        assert backing.saved == {
            "mapping": {"people/a": "graph-a"},
            "hashes": {"people/a": "hash-a"},
            "sync_token": "token-2",
        }

    async def test_remove_deletes_backing_file(self, monkeypatch):
        s, backing = make_store(monkeypatch)
        await s.async_remove()
        assert backing.removed is True


class TestMappingInvariant:
    async def test_conflicting_mapping_is_rejected(self, monkeypatch):
        s, _ = make_store(monkeypatch)
        await s.async_load()
        s.save_mapping("people/a", "outlook-1", "hash-a")

        with pytest.raises(MappingConflictError, match="outlook-1"):
            s.save_mapping("people/b", "outlook-1", "hash-b")

        assert s.lookup("people/a") == "outlook-1"
        assert s.lookup("people/b") is None

    async def test_idempotent_resave_is_allowed(self, monkeypatch):
        s, _ = make_store(monkeypatch)
        await s.async_load()
        s.save_mapping("people/a", "outlook-1", "hash-a")
        s.save_mapping("people/a", "outlook-1", "hash-a2")

        assert s.lookup("people/a") == "outlook-1"
        assert s.get_hash("people/a") == "hash-a2"

    async def test_resave_allowed_despite_preexisting_conflict(self, monkeypatch):
        """A store corrupted by an older version must stay updatable."""
        stored = {"mapping": {"people/a": "outlook-1", "people/b": "outlook-1"}}
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        s.save_mapping("people/a", "outlook-1", "new-hash")
        assert s.get_hash("people/a") == "new-hash"

    async def test_remapping_a_resource_to_a_free_id_is_allowed(self, monkeypatch):
        s, _ = make_store(monkeypatch)
        await s.async_load()
        s.save_mapping("people/a", "outlook-1", "hash-a")
        s.save_mapping("people/a", "outlook-2", "hash-a")

        assert s.lookup("people/a") == "outlook-2"

    async def test_resource_names_for_contact_id(self, monkeypatch):
        stored = {"mapping": {"people/a": "outlook-1", "people/b": "outlook-2"}}
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        assert s.resource_names_for_contact_id("outlook-1") == {"people/a"}
        assert s.resource_names_for_contact_id("missing") == set()

    async def test_find_mapping_conflicts_detects_many_to_one(self, monkeypatch):
        stored = {
            "mapping": {
                "people/a": "outlook-1",
                "people/b": "outlook-1",
                "people/c": "outlook-2",
            }
        }
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        assert s.find_mapping_conflicts() == {"outlook-1": {"people/a", "people/b"}}

    async def test_find_mapping_conflicts_empty_when_clean(self, monkeypatch):
        stored = {"mapping": {"people/a": "outlook-1", "people/b": "outlook-2"}}
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        assert s.find_mapping_conflicts() == {}


class TestMappingOperations:
    async def test_delete_mapping_removes_hash_too(self, monkeypatch):
        s, _ = make_store(monkeypatch)
        await s.async_load()
        s.save_mapping("people/a", "graph-a", "hash-a")
        s.delete_mapping("people/a")

        assert s.lookup("people/a") is None
        assert s.get_hash("people/a") is None

    async def test_delete_unknown_mapping_is_a_noop(self, monkeypatch):
        s, _ = make_store(monkeypatch)
        await s.async_load()
        s.delete_mapping("people/nope")
        assert s.all_resource_names() == set()

    async def test_repoint_moves_every_affected_resource(self, monkeypatch):
        stored = {
            "mapping": {
                "people/a": "outlook-old",
                "people/b": "outlook-other",
            }
        }
        s, _ = make_store(monkeypatch, stored)
        await s.async_load()

        s.repoint_mapping("outlook-old", "outlook-new")

        assert s.lookup("people/a") == "outlook-new"
        assert s.lookup("people/b") == "outlook-other"
