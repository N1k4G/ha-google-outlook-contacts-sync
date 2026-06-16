"""Tests for sync/engine.py — all external calls mocked."""

from __future__ import annotations

from typing import Any
from unittest.mock import ANY, AsyncMock, MagicMock

import pytest

from custom_components.google_outlook_contacts_sync.sync.engine import (
    SyncEngine,
    SyncPlan,
    SyncResult,
)
from custom_components.google_outlook_contacts_sync.sync.google_client import (
    SyncTokenExpiredError,
)


def _make_engine(
    hass: Any,
    people: list[dict[str, Any]],
    sync_token: str | None = "token_new",
    existing_mapping: dict[str, str] | None = None,
    existing_hashes: dict[str, str] | None = None,
    delete_removed: bool = False,
    auto_remove_duplicates: bool = False,
    outlook_raw: list[dict[str, Any]] | None = None,
) -> tuple[SyncEngine, MagicMock, MagicMock, MagicMock]:
    """Build a SyncEngine with mocked clients and store."""
    google_client = MagicMock()
    google_client.fetch_contacts.return_value = (people, sync_token)

    graph_client = MagicMock()
    graph_client.create_contact.return_value = "graph-id-001"
    graph_client.update_contact.return_value = None
    graph_client.delete_contact.return_value = None
    graph_client.list_contacts_raw.return_value = outlook_raw or []

    store = MagicMock()
    store.sync_token = None
    store.async_save = AsyncMock()
    mapping = existing_mapping or {}
    hashes = existing_hashes or {}
    store.lookup.side_effect = lambda r: mapping.get(r)
    store.get_hash.side_effect = lambda r: hashes.get(r)
    store.save_mapping.side_effect = lambda r, cid, h: (
        mapping.update({r: cid}),
        hashes.update({r: h}),
    )
    store.delete_mapping.side_effect = lambda r: (
        mapping.pop(r, None),
        hashes.pop(r, None),
    )
    store.all_resource_names.return_value = set(mapping.keys())
    store.mapped_ids.side_effect = lambda: set(mapping.values())

    def _repoint(old: str, new: str) -> None:
        for rn, cid in list(mapping.items()):
            if cid == old:
                mapping[rn] = new

    store.repoint_mapping.side_effect = _repoint

    hass.async_add_executor_job = AsyncMock(
        side_effect=lambda fn, *args: fn(*args)
    )

    engine = SyncEngine(
        hass=hass,
        google_client=google_client,
        graph_client=graph_client,
        store=store,
        delete_removed=delete_removed,
        auto_remove_duplicates=auto_remove_duplicates,
    )
    return engine, google_client, graph_client, store


@pytest.fixture
def hass() -> MagicMock:
    return MagicMock()


@pytest.fixture
def simple_person() -> dict[str, Any]:
    return {
        "resourceName": "people/c123",
        "names": [
            {
                "givenName": "Alice",
                "familyName": "Smith",
                "displayName": "Alice Smith",
            }
        ],
        "emailAddresses": [{"value": "alice@example.com", "type": "home"}],
    }


class TestSyncEngineCreate:
    @pytest.mark.asyncio
    async def test_creates_new_contact(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(hass, [simple_person])
        result = await engine.async_sync()

        graph.create_contact.assert_called_once()
        store.save_mapping.assert_called_once()
        assert result.created == 1
        assert result.updated == 0
        assert result.failed == 0

    @pytest.mark.asyncio
    async def test_stores_new_sync_token(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, _, store = _make_engine(hass, [simple_person], sync_token="tok_123")
        await engine.async_sync()

        assert store.sync_token == "tok_123"
        store.async_save.assert_called_once()


class TestSyncEngineUpdate:
    @pytest.mark.asyncio
    async def test_skips_unchanged_contact(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.mapping import (
            contact_hash,
            to_graph_contact,
        )

        graph_contact = to_graph_contact(simple_person)
        h = contact_hash(graph_contact)

        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "graph-id-existing"},
            existing_hashes={"people/c123": h},
        )
        result = await engine.async_sync()

        graph.update_contact.assert_not_called()
        assert result.skipped == 1
        assert result.updated == 0

    @pytest.mark.asyncio
    async def test_updates_changed_contact(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "graph-id-existing"},
            existing_hashes={"people/c123": "stale_hash"},
        )
        result = await engine.async_sync()

        graph.update_contact.assert_called_once_with("graph-id-existing", ANY)
        assert result.updated == 1


class TestSyncEngineDelete:
    @pytest.mark.asyncio
    async def test_does_not_delete_when_disabled(self, hass: MagicMock) -> None:
        deleted_person: dict[str, Any] = {
            "resourceName": "people/c999",
            "metadata": {"deleted": True},
        }
        engine, _, graph, store = _make_engine(
            hass,
            [deleted_person],
            existing_mapping={"people/c999": "graph-id-del"},
            delete_removed=False,
        )
        result = await engine.async_sync()

        graph.delete_contact.assert_not_called()
        assert result.skipped == 1

    @pytest.mark.asyncio
    async def test_deletes_when_enabled(self, hass: MagicMock) -> None:
        deleted_person: dict[str, Any] = {
            "resourceName": "people/c999",
            "metadata": {"deleted": True},
        }
        engine, _, graph, store = _make_engine(
            hass,
            [deleted_person],
            existing_mapping={"people/c999": "graph-id-del"},
            delete_removed=True,
        )
        result = await engine.async_sync()

        graph.delete_contact.assert_called_once_with("graph-id-del")
        assert result.deleted == 1

    @pytest.mark.asyncio
    async def test_deleted_person_not_in_mapping_skipped(
        self, hass: MagicMock
    ) -> None:
        deleted_person: dict[str, Any] = {
            "resourceName": "people/c888",
            "metadata": {"deleted": True},
        }
        engine, _, graph, _ = _make_engine(
            hass, [deleted_person], delete_removed=True
        )
        result = await engine.async_sync()

        graph.delete_contact.assert_not_called()
        assert result.skipped == 1


class TestSyncEngineErrors:
    @pytest.mark.asyncio
    async def test_per_contact_error_does_not_abort_run(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        good_person: dict[str, Any] = {
            "resourceName": "people/c200",
            "names": [{"givenName": "Bob", "displayName": "Bob"}],
        }
        engine, _, graph, _ = _make_engine(hass, [simple_person, good_person])
        graph.create_contact.side_effect = [RuntimeError("Graph error"), "graph-id-002"]

        result = await engine.async_sync()

        assert result.failed == 1
        assert result.created == 1
        assert len(result.errors) == 1

    @pytest.mark.asyncio
    async def test_google_fetch_failure_returns_result(
        self, hass: MagicMock
    ) -> None:
        google_client = MagicMock()
        google_client.fetch_contacts.side_effect = ConnectionError("Network down")
        graph_client = MagicMock()
        store = MagicMock()
        store.sync_token = None
        store.async_save = AsyncMock()

        hass.async_add_executor_job = AsyncMock(
            side_effect=lambda fn, *args: fn(*args)
        )

        engine = SyncEngine(hass, google_client, graph_client, store)
        result = await engine.async_sync()

        assert result.failed == 1
        assert "Google fetch failed" in result.errors[0]

    @pytest.mark.asyncio
    async def test_expired_sync_token_triggers_full_resync(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        google_client = MagicMock()
        google_client.fetch_contacts.side_effect = [
            SyncTokenExpiredError("expired"),
            ([simple_person], "new_token"),
        ]
        graph_client = MagicMock()
        graph_client.create_contact.return_value = "id-001"
        store = MagicMock()
        store.sync_token = "old_token"
        store.async_save = AsyncMock()
        store.lookup.return_value = None
        store.get_hash.return_value = None
        store.save_mapping = MagicMock()

        hass.async_add_executor_job = AsyncMock(
            side_effect=lambda fn, *args: fn(*args)
        )

        engine = SyncEngine(hass, google_client, graph_client, store)
        result = await engine.async_sync()

        assert google_client.fetch_contacts.call_count == 2
        assert result.created == 1


class TestSyncResult:
    def test_as_dict_contains_all_fields(self) -> None:
        r = SyncResult(created=1, updated=2, deleted=3, skipped=4, failed=5)
        d = r.as_dict()
        assert d["created"] == 1
        assert d["updated"] == 2
        assert d["deleted"] == 3
        assert d["skipped"] == 4
        assert d["failed"] == 5
        assert "errors" in d

    def test_summary_includes_failures_only_when_present(self) -> None:
        assert "failed" not in SyncResult(created=1).summary()
        assert "failed" in SyncResult(failed=2).summary()


class TestDryRun:
    @pytest.mark.asyncio
    async def test_plan_classifies_create_update_skip(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.mapping import (
            contact_hash,
            to_graph_contact,
        )

        # One brand-new person, one already-synced with a stale hash (update),
        # one already-synced and unchanged (skip).
        changed = {
            "resourceName": "people/c200",
            "names": [{"givenName": "Bob", "displayName": "Bob"}],
        }
        unchanged = {
            "resourceName": "people/c300",
            "names": [{"givenName": "Eve", "displayName": "Eve"}],
        }
        unchanged_hash = contact_hash(to_graph_contact(unchanged))

        engine, _, graph, store = _make_engine(
            hass,
            [simple_person, changed, unchanged],
            existing_mapping={
                "people/c200": "graph-200",
                "people/c300": "graph-300",
            },
            existing_hashes={
                "people/c200": "stale",
                "people/c300": unchanged_hash,
            },
        )

        plan = await engine.async_plan()

        assert isinstance(plan, SyncPlan)
        assert {c.resource_name for c in plan.to_create} == {"people/c123"}
        assert {c.resource_name for c in plan.to_update} == {"people/c200"}
        assert plan.to_delete == []
        assert plan.total == 2
        # Dry run must change nothing and persist nothing.
        graph.create_contact.assert_not_called()
        graph.update_contact.assert_not_called()
        graph.delete_contact.assert_not_called()
        store.save_mapping.assert_not_called()
        store.async_save.assert_not_called()

    @pytest.mark.asyncio
    async def test_plan_lists_delete_when_enabled(self, hass: MagicMock) -> None:
        deleted_person: dict[str, Any] = {
            "resourceName": "people/c999",
            "metadata": {"deleted": True},
        }
        engine, _, _, _ = _make_engine(
            hass,
            [deleted_person],
            existing_mapping={"people/c999": "graph-del"},
            delete_removed=True,
        )

        plan = await engine.async_plan()

        assert {c.resource_name for c in plan.to_delete} == {"people/c999"}

    @pytest.mark.asyncio
    async def test_plan_does_not_clear_sync_token_on_expiry(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        google_client = MagicMock()
        google_client.fetch_contacts.side_effect = [
            SyncTokenExpiredError("expired"),
            ([simple_person], "ignored"),
        ]
        store = MagicMock()
        store.sync_token = "keep_me"
        store.lookup.return_value = None
        store.get_hash.return_value = None

        hass.async_add_executor_job = AsyncMock(
            side_effect=lambda fn, *args: fn(*args)
        )

        engine = SyncEngine(hass, google_client, MagicMock(), store)
        plan = await engine.async_plan()

        # Token preserved (dry run must not disturb real-sync state).
        assert store.sync_token == "keep_me"
        assert google_client.fetch_contacts.call_count == 2
        assert plan.total == 1


class TestSyncTokenOnFailure:
    @pytest.mark.asyncio
    async def test_token_not_advanced_when_a_contact_fails(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass, [simple_person], sync_token="tok_after"
        )
        graph.create_contact.side_effect = RuntimeError("boom")

        result = await engine.async_sync()

        assert result.failed == 1
        # Old token (None) preserved so the failed contact is retried next run.
        assert store.sync_token is None
        store.async_save.assert_called_once()

    @pytest.mark.asyncio
    async def test_token_advanced_on_clean_run(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, _, store = _make_engine(
            hass, [simple_person], sync_token="tok_after"
        )
        await engine.async_sync()
        assert store.sync_token == "tok_after"


class TestFullSync:
    @pytest.mark.asyncio
    async def test_adopts_existing_outlook_contact_by_name(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        # Google "Alice Smith" is unmapped; an Outlook "Alice Smith" exists.
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            outlook_raw=[{"id": "outlook-9", "displayName": "Alice Smith"}],
        )

        result = await engine.async_full_sync()

        # Adopted, not created: update the existing Outlook contact.
        graph.create_contact.assert_not_called()
        graph.update_contact.assert_called_once_with("outlook-9", ANY)
        assert result.created == 0
        assert result.updated == 1
        store.save_mapping.assert_called_once_with("people/c123", "outlook-9", ANY)

    @pytest.mark.asyncio
    async def test_creates_when_no_outlook_match(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(hass, [simple_person], outlook_raw=[])
        result = await engine.async_full_sync()

        graph.create_contact.assert_called_once()
        assert result.created == 1


class TestDedup:
    @pytest.mark.asyncio
    async def test_removes_duplicate_keeping_most_complete(
        self, hass: MagicMock
    ) -> None:
        # Two Outlook contacts share a name; outlook-1 is more complete.
        outlook_raw = [
            {
                "id": "outlook-1",
                "displayName": "Bob Jones",
                "givenName": "Bob",
                "surname": "Jones",
                "emailAddresses": [{"address": "b@x.com"}],
            },
            {"id": "outlook-2", "displayName": "Bob Jones"},
        ]
        engine, _, graph, store = _make_engine(
            hass,
            people=[],
            auto_remove_duplicates=True,
            outlook_raw=outlook_raw,
        )

        result = await engine.async_sync()

        graph.delete_contact.assert_called_once_with("outlook-2")
        store.repoint_mapping.assert_called_once_with("outlook-2", "outlook-1")
        assert result.duplicates_removed == 1
        assert "bob jones" in result.duplicates

    @pytest.mark.asyncio
    async def test_no_dedup_when_disabled(self, hass: MagicMock) -> None:
        outlook_raw = [
            {"id": "o1", "displayName": "Bob Jones"},
            {"id": "o2", "displayName": "Bob Jones"},
        ]
        engine, _, graph, _ = _make_engine(
            hass, people=[], auto_remove_duplicates=False, outlook_raw=outlook_raw
        )
        result = await engine.async_sync()

        graph.delete_contact.assert_not_called()
        assert result.duplicates_removed == 0
