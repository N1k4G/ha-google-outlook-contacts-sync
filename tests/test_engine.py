"""Tests for sync/engine.py — all external calls mocked."""

from __future__ import annotations

import asyncio
from typing import Any
from unittest.mock import ANY, AsyncMock, MagicMock, patch

import pytest

from custom_components.google_outlook_contacts_sync.auth import CredentialStatus
from custom_components.google_outlook_contacts_sync.sync import (
    graph_client as graph_client_module,
)
from custom_components.google_outlook_contacts_sync.sync.engine import (
    ABORT_AFTER_FAILURES,
    MAX_REPORTED_ERRORS,
    SyncEngine,
    SyncPlan,
    SyncResult,
)
from custom_components.google_outlook_contacts_sync.sync.google_client import (
    SyncTokenExpiredError,
)
from custom_components.google_outlook_contacts_sync.sync.graph_client import (
    GraphContactsClient,
    MSAuthError,
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
    stored_sync_token: str | None = "token_stored",
) -> tuple[SyncEngine, MagicMock, MagicMock, MagicMock]:
    """Build a SyncEngine with mocked clients and store.

    *sync_token* is what the Google fetch returns; *stored_sync_token* is the
    cursor already in the store. The latter defaults to a real token because
    async_sync() routes to full reconciliation when it is None (#33), so a
    delta-path test must start with a cursor.
    """
    google_client = MagicMock()
    google_client.fetch_contacts.return_value = (people, sync_token)

    graph_client = MagicMock()
    graph_client.create_contact.return_value = "graph-id-001"
    graph_client.update_contact.return_value = None
    graph_client.delete_contact.return_value = None
    graph_client.list_contacts_raw.return_value = outlook_raw or []

    store = MagicMock()
    store.sync_token = stored_sync_token
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

    hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *args: fn(*args))

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
            to_graph_contact_update,
        )

        # The engine hashes the authoritative state, not the sparse create body.
        h = contact_hash(to_graph_contact_update(simple_person))

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
    async def test_deleted_person_not_in_mapping_skipped(self, hass: MagicMock) -> None:
        deleted_person: dict[str, Any] = {
            "resourceName": "people/c888",
            "metadata": {"deleted": True},
        }
        engine, _, graph, _ = _make_engine(hass, [deleted_person], delete_removed=True)
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
    async def test_google_fetch_failure_propagates(self, hass: MagicMock) -> None:
        """#14: a failed fetch must not look like a completed run.

        It used to return a SyncResult with finished_at set, which the
        coordinator accepted as success -- the last-sync sensor jumped to the
        failure time and only the problem sensor showed anything was wrong.
        """
        google_client = MagicMock()
        google_client.fetch_contacts.side_effect = ConnectionError("Network down")
        graph_client = MagicMock()
        store = MagicMock()
        store.sync_token = "token_stored"
        store.async_save = AsyncMock()

        hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *args: fn(*args))

        engine = SyncEngine(hass, google_client, graph_client, store)
        with pytest.raises(ConnectionError):
            await engine.async_sync()

        store.async_save.assert_not_called()

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
        graph_client.list_contacts_raw.return_value = []
        store = MagicMock()
        store.sync_token = "old_token"
        store.async_save = AsyncMock()
        store.lookup.return_value = None
        store.get_hash.return_value = None
        store.save_mapping = MagicMock()

        hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *args: fn(*args))

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

    def test_summary_marks_an_aborted_run(self) -> None:
        assert "aborted" not in SyncResult(failed=2).summary()
        assert SyncResult(failed=10, aborted=True).summary().endswith("aborted")

    def test_record_failure_counts_all_but_keeps_only_the_report_limit(self) -> None:
        r = SyncResult()
        for i in range(MAX_REPORTED_ERRORS + 5):
            r.record_failure(f"people/c{i}: boom")
        assert r.failed == MAX_REPORTED_ERRORS + 5
        assert len(r.errors) == MAX_REPORTED_ERRORS
        assert r.errors[0] == "people/c0: boom"


class TestFieldClearing:
    """Regressions for #24: removing a field in Google must clear it in Outlook."""

    @staticmethod
    def _person(**extra: Any) -> dict[str, Any]:
        return {
            "resourceName": "people/c123",
            "names": [{"givenName": "Alice", "displayName": "Alice Smith"}],
            **extra,
        }

    @pytest.mark.asyncio
    async def test_removed_field_is_explicitly_cleared(self, hass: MagicMock) -> None:
        from custom_components.google_outlook_contacts_sync.sync.mapping import (
            contact_hash,
            to_graph_contact_update,
        )

        was = self._person(
            phoneNumbers=[{"value": "+49 170 1", "type": "mobile"}],
            birthdays=[
                {"metadata": {"primary": True}, "date": {"month": 3, "day": 15}}
            ],
        )
        now = self._person()

        engine, _, graph, store = _make_engine(
            hass,
            [now],
            existing_mapping={"people/c123": "graph-1"},
            existing_hashes={"people/c123": contact_hash(to_graph_contact_update(was))},
        )
        result = await engine.async_sync()

        graph.update_contact.assert_called_once()
        _, body = graph.update_contact.call_args[0]
        assert body["mobilePhone"] is None
        assert body["birthday"] is None
        assert body["homePhones"] == []
        assert result.updated == 1
        assert store.get_hash("people/c123") == contact_hash(
            to_graph_contact_update(now)
        )

    @pytest.mark.asyncio
    async def test_second_run_with_stored_hash_skips(self, hass: MagicMock) -> None:
        from custom_components.google_outlook_contacts_sync.sync.mapping import (
            contact_hash,
            to_graph_contact_update,
        )

        person = self._person()
        engine, _, graph, _ = _make_engine(
            hass,
            [person],
            existing_mapping={"people/c123": "graph-1"},
            existing_hashes={
                "people/c123": contact_hash(to_graph_contact_update(person))
            },
        )
        result = await engine.async_sync()

        graph.update_contact.assert_not_called()
        assert result.skipped == 1

    @pytest.mark.asyncio
    async def test_mapped_contact_stripped_of_all_data_still_updates(
        self, hass: MagicMock
    ) -> None:
        """Previously skipped outright, leaving stale Outlook values forever."""
        emptied = {"resourceName": "people/c123"}

        engine, _, graph, _ = _make_engine(
            hass,
            [emptied],
            existing_mapping={"people/c123": "graph-1"},
            existing_hashes={"people/c123": "some-old-hash"},
        )
        result = await engine.async_sync()

        graph.update_contact.assert_called_once()
        _, body = graph.update_contact.call_args[0]
        assert body["givenName"] is None
        assert body["emailAddresses"] == []
        assert result.updated == 1

    @pytest.mark.asyncio
    async def test_unmapped_contact_with_no_data_is_not_created(
        self, hass: MagicMock
    ) -> None:
        engine, _, graph, _ = _make_engine(hass, [{"resourceName": "people/empty"}])
        result = await engine.async_sync()

        graph.create_contact.assert_not_called()
        assert result.skipped == 1

    @pytest.mark.asyncio
    async def test_create_sends_sparse_body(self, hass: MagicMock) -> None:
        engine, _, graph, _ = _make_engine(hass, [self._person()])
        await engine.async_sync()

        (body,) = graph.create_contact.call_args[0]
        assert "mobilePhone" not in body
        assert "otherAddress" not in body


class TestDryRun:
    @pytest.mark.asyncio
    async def test_plan_classifies_create_update_skip(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.mapping import (
            contact_hash,
            to_graph_contact_update,
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
        unchanged_hash = contact_hash(to_graph_contact_update(unchanged))

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

        hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *args: fn(*args))

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
        # Previous cursor preserved so the failed contact is retried next run.
        assert store.sync_token == "token_stored"
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
    async def test_adopts_on_unique_email_match(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """Case matters neither way — comparison is strip()+casefold()."""
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            outlook_raw=[
                {
                    "id": "outlook-9",
                    "displayName": "Completely Different Name",
                    "emailAddresses": [{"address": "  ALICE@example.com "}],
                }
            ],
        )

        result = await engine.async_full_sync()

        graph.create_contact.assert_not_called()
        graph.update_contact.assert_called_once_with("outlook-9", ANY)
        assert result.updated == 1
        store.save_mapping.assert_called_once_with("people/c123", "outlook-9", ANY)

    @pytest.mark.asyncio
    async def test_same_name_different_email_is_not_adopted(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """Regression for #25: name-only adoption overwrote unrelated people."""
        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            outlook_raw=[
                {
                    "id": "outlook-9",
                    "displayName": "Alice Smith",
                    "emailAddresses": [{"address": "other.alice@example.com"}],
                }
            ],
        )

        result = await engine.async_full_sync()

        graph.update_contact.assert_not_called()
        graph.create_contact.assert_called_once()
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_same_name_but_google_has_no_email_is_not_adopted(
        self, hass: MagicMock
    ) -> None:
        person = {
            "resourceName": "people/c123",
            "names": [{"displayName": "Alice Smith"}],
        }
        engine, _, graph, _ = _make_engine(
            hass,
            [person],
            outlook_raw=[
                {
                    "id": "outlook-9",
                    "displayName": "Alice Smith",
                    "emailAddresses": [{"address": "alice@example.com"}],
                }
            ],
        )

        result = await engine.async_full_sync()

        graph.update_contact.assert_not_called()
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_ambiguous_email_is_not_adopted(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            outlook_raw=[
                {
                    "id": "outlook-1",
                    "emailAddresses": [{"address": "alice@example.com"}],
                },
                {
                    "id": "outlook-2",
                    "emailAddresses": [{"address": "alice@example.com"}],
                },
            ],
        )

        result = await engine.async_full_sync()

        graph.update_contact.assert_not_called()
        graph.create_contact.assert_called_once()
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_already_mapped_outlook_contact_is_not_a_candidate(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/other": "outlook-9"},
            outlook_raw=[
                {
                    "id": "outlook-9",
                    "emailAddresses": [{"address": "alice@example.com"}],
                }
            ],
        )

        result = await engine.async_full_sync()

        graph.create_contact.assert_called_once()
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_match_on_secondary_email_claims_the_contact_once(
        self, hass: MagicMock
    ) -> None:
        """A claimed id must leave every email bucket, not just the matched one."""
        first = {
            "resourceName": "people/c1",
            "names": [{"displayName": "One"}],
            "emailAddresses": [{"value": "second@example.com"}],
        }
        second = {
            "resourceName": "people/c2",
            "names": [{"displayName": "Two"}],
            "emailAddresses": [{"value": "first@example.com"}],
        }
        engine, _, graph, _ = _make_engine(
            hass,
            [first, second],
            outlook_raw=[
                {
                    "id": "outlook-9",
                    "emailAddresses": [
                        {"address": "first@example.com"},
                        {"address": "second@example.com"},
                    ],
                }
            ],
        )

        result = await engine.async_full_sync()

        graph.update_contact.assert_called_once_with("outlook-9", ANY)
        assert result.updated == 1
        assert result.created == 1  # the second person gets a fresh contact

    @pytest.mark.asyncio
    async def test_creates_when_no_outlook_match(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(hass, [simple_person], outlook_raw=[])
        result = await engine.async_full_sync()

        graph.create_contact.assert_called_once()
        assert result.created == 1


class TestStaleMappingRepair:
    """#30: a mapped Outlook contact deleted by hand must be recreated."""

    @pytest.mark.asyncio
    async def test_missing_mapped_id_leads_to_create(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-missing"},
            outlook_raw=[{"id": "outlook-unrelated"}],
        )

        result = await engine.async_full_sync()

        store.delete_mapping.assert_any_call("people/c123")
        graph.update_contact.assert_not_called()
        graph.create_contact.assert_called_once()
        store.save_mapping.assert_called_with("people/c123", "graph-id-001", ANY)
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_missing_mapped_id_can_adopt_a_unique_match(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-missing"},
            outlook_raw=[
                {
                    "id": "outlook-new",
                    "emailAddresses": [{"address": "alice@example.com"}],
                }
            ],
        )

        result = await engine.async_full_sync()

        store.delete_mapping.assert_any_call("people/c123")
        graph.create_contact.assert_not_called()
        graph.update_contact.assert_called_once_with("outlook-new", ANY)
        assert result.updated == 1

    @pytest.mark.asyncio
    async def test_existing_mapped_id_keeps_its_mapping(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1"},
            outlook_raw=[{"id": "outlook-1"}],
        )

        result = await engine.async_full_sync()

        store.delete_mapping.assert_not_called()
        graph.create_contact.assert_not_called()
        graph.update_contact.assert_called_once_with("outlook-1", ANY)
        assert result.updated == 1

    @pytest.mark.asyncio
    async def test_snapshot_failure_removes_no_mappings(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """A partial listing must never be read as 'these contacts are gone'."""
        from custom_components.google_outlook_contacts_sync.sync.graph_client import (
            GraphRequestError,
        )

        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1"},
        )
        graph.list_contacts_raw.side_effect = GraphRequestError("listing failed")

        with pytest.raises(GraphRequestError):
            await engine.async_full_sync()

        store.delete_mapping.assert_not_called()
        graph.create_contact.assert_not_called()

    @pytest.mark.asyncio
    async def test_stale_mapping_stays_removed_when_create_fails(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """A provably dead Outlook id is no use as a retry anchor."""
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-missing"},
            outlook_raw=[],
        )
        graph.create_contact.side_effect = RuntimeError("boom")

        result = await engine.async_full_sync()

        store.delete_mapping.assert_any_call("people/c123")
        assert store.lookup("people/c123") is None
        assert result.failed == 1
        assert result.created == 0


class TestOrphanReconciliation:
    """#18: mappings whose Google contact is gone must not linger."""

    @pytest.mark.asyncio
    async def test_delta_delete_untracks_when_option_off(self, hass: MagicMock) -> None:
        deleted = {"resourceName": "people/c999", "metadata": {"deleted": True}}
        engine, _, graph, store = _make_engine(
            hass,
            [deleted],
            existing_mapping={"people/c999": "graph-del"},
            delete_removed=False,
        )

        result = await engine.async_sync()

        graph.delete_contact.assert_not_called()
        store.delete_mapping.assert_called_once_with("people/c999")
        assert result.deleted == 0

    @pytest.mark.asyncio
    async def test_delta_delete_removes_then_untracks_when_option_on(
        self, hass: MagicMock
    ) -> None:
        deleted = {"resourceName": "people/c999", "metadata": {"deleted": True}}
        engine, _, graph, store = _make_engine(
            hass,
            [deleted],
            existing_mapping={"people/c999": "graph-del"},
            delete_removed=True,
        )

        result = await engine.async_sync()

        graph.delete_contact.assert_called_once_with("graph-del")
        store.delete_mapping.assert_called_once_with("people/c999")
        assert result.deleted == 1

    @pytest.mark.asyncio
    async def test_full_resync_orphan_untracked_when_option_off(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1", "people/old": "graph-old"},
            outlook_raw=[{"id": "outlook-1"}, {"id": "graph-old"}],
            delete_removed=False,
        )

        result = await engine.async_full_sync()

        graph.delete_contact.assert_not_called()
        store.delete_mapping.assert_any_call("people/old")
        assert result.deleted == 0

    @pytest.mark.asyncio
    async def test_full_resync_orphan_deleted_when_option_on(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1", "people/old": "graph-old"},
            outlook_raw=[{"id": "outlook-1"}, {"id": "graph-old"}],
            delete_removed=True,
        )

        result = await engine.async_full_sync()

        graph.delete_contact.assert_called_once_with("graph-old")
        store.delete_mapping.assert_any_call("people/old")
        assert result.deleted == 1

    @pytest.mark.asyncio
    async def test_failed_orphan_delete_keeps_mapping_for_retry(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1", "people/old": "graph-old"},
            outlook_raw=[{"id": "outlook-1"}, {"id": "graph-old"}],
            delete_removed=True,
        )
        graph.delete_contact.side_effect = RuntimeError("boom")

        result = await engine.async_full_sync()

        assert result.failed == 1
        assert result.deleted == 0
        assert store.lookup("people/old") == "graph-old"
        # failed > 0 invalidates the cursor, forcing another full pass.
        assert store.sync_token is None

    @pytest.mark.asyncio
    async def test_active_contact_without_mappable_fields_is_not_an_orphan(
        self, hass: MagicMock
    ) -> None:
        """Present but empty is not the same as deleted."""
        empty = {"resourceName": "people/empty"}
        engine, _, graph, store = _make_engine(
            hass,
            [empty],
            existing_mapping={"people/empty": "outlook-1"},
            outlook_raw=[{"id": "outlook-1"}],
            delete_removed=True,
        )

        await engine.async_full_sync()

        graph.delete_contact.assert_not_called()
        assert store.lookup("people/empty") == "outlook-1"

    @pytest.mark.asyncio
    async def test_auth_error_during_orphan_delete_propagates(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.graph_client import (
            MSAuthError,
        )

        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/c123": "outlook-1", "people/old": "graph-old"},
            outlook_raw=[{"id": "outlook-1"}, {"id": "graph-old"}],
            delete_removed=True,
        )
        graph.delete_contact.side_effect = MSAuthError("token gone")

        with pytest.raises(MSAuthError):
            await engine.async_full_sync()


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


def _first_call_gate(target: Any):
    """Executor that suspends the *first* call to one target function.

    Only the first call blocks, so a second concurrent run is held up purely by
    the engine lock — never by the fixture. Without that distinction these tests
    pass whether or not the lock exists.

    Returned as a module-level async function rather than a callable object:
    AsyncMock only awaits a side_effect that asyncio.iscoroutinefunction()
    accepts, which is False for an instance with an async __call__.
    """
    reached = asyncio.Event()
    release = asyncio.Event()
    armed = [True]

    async def executor(fn: Any, *args: Any) -> Any:
        if armed[0] and fn is target:
            armed[0] = False
            reached.set()
            await release.wait()
        return fn(*args)

    executor.reached = reached  # type: ignore[attr-defined]
    executor.release = release  # type: ignore[attr-defined]
    return executor


async def _settle(iterations: int = 20) -> None:
    """Let every runnable task progress as far as it can."""
    for _ in range(iterations):
        await asyncio.sleep(0)


def _dupe(contact_id: str, name: str = "Bob Jones", email: str | None = None):
    raw: dict[str, Any] = {"id": contact_id, "displayName": name}
    if email:
        raw["emailAddresses"] = [{"address": email}]
    return raw


class TestDedupSafety:
    """#27 error handling and #31 many-to-one protection."""

    @pytest.mark.asyncio
    async def test_auth_error_while_listing_propagates(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.graph_client import (
            MSAuthError,
        )

        engine, _, graph, _ = _make_engine(
            hass, [simple_person], auto_remove_duplicates=True
        )
        graph.list_contacts_raw.side_effect = MSAuthError("token gone")

        with pytest.raises(MSAuthError):
            await engine.async_sync()

    @pytest.mark.asyncio
    async def test_auth_error_while_deleting_propagates(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        from custom_components.google_outlook_contacts_sync.sync.graph_client import (
            MSAuthError,
        )

        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            auto_remove_duplicates=True,
            outlook_raw=[_dupe("outlook-1"), _dupe("outlook-2")],
        )
        graph.delete_contact.side_effect = MSAuthError("token gone")

        with pytest.raises(MSAuthError):
            await engine.async_sync()

        store.repoint_mapping.assert_not_called()

    @pytest.mark.asyncio
    async def test_listing_failure_counts_as_failed(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """errors alone left the problem sensor off and advanced the cursor."""
        engine, _, graph, store = _make_engine(
            hass, [simple_person], auto_remove_duplicates=True, sync_token="tok_after"
        )
        graph.list_contacts_raw.side_effect = RuntimeError("graph down")

        result = await engine.async_sync()

        assert result.failed == 1
        assert any("Dedup listing failed" in e for e in result.errors)
        assert store.sync_token == "token_stored"

    @pytest.mark.asyncio
    async def test_delete_failure_counts_and_processing_continues(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            auto_remove_duplicates=True,
            outlook_raw=[
                _dupe("outlook-keep"),
                _dupe("outlook-bad"),
                _dupe("outlook-good"),
            ],
        )

        def _delete(contact_id: str) -> None:
            if contact_id == "outlook-bad":
                raise RuntimeError("boom")

        graph.delete_contact.side_effect = _delete

        result = await engine.async_sync()

        assert result.failed == 1
        assert result.duplicates_removed == 1
        assert graph.delete_contact.call_count == 2

    @pytest.mark.asyncio
    async def test_group_with_two_tracked_contacts_is_skipped(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """Regression for #31: merging these corrupts the mapping permanently."""
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/a": "outlook-a", "people/b": "outlook-b"},
            auto_remove_duplicates=True,
            outlook_raw=[_dupe("outlook-a"), _dupe("outlook-b")],
        )

        result = await engine.async_sync()

        graph.delete_contact.assert_not_called()
        store.repoint_mapping.assert_not_called()
        assert result.duplicates_removed == 0

    @pytest.mark.asyncio
    async def test_one_tracked_contact_becomes_the_keeper(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass,
            [simple_person],
            existing_mapping={"people/a": "outlook-a"},
            auto_remove_duplicates=True,
            outlook_raw=[_dupe("outlook-a"), _dupe("outlook-untracked")],
        )

        result = await engine.async_sync()

        graph.delete_contact.assert_called_once_with("outlook-untracked")
        assert result.duplicates_removed == 1
        assert store.lookup("people/a") == "outlook-a"

    @pytest.mark.asyncio
    async def test_untracked_group_still_deduplicates(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, _ = _make_engine(
            hass,
            [simple_person],
            auto_remove_duplicates=True,
            outlook_raw=[_dupe("outlook-1"), _dupe("outlook-2")],
        )

        result = await engine.async_sync()

        assert graph.delete_contact.call_count == 1
        assert result.duplicates_removed == 1

    @pytest.mark.asyncio
    async def test_duplicates_records_one_entry_per_group(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, _, _ = _make_engine(
            hass,
            [simple_person],
            auto_remove_duplicates=True,
            outlook_raw=[_dupe("o-1"), _dupe("o-2"), _dupe("o-3")],
        )

        result = await engine.async_sync()

        assert result.duplicates_removed == 2
        assert result.duplicates == ["bob jones"]


class TestRunSerialization:
    """#8: every public entry point shares one engine-local lock."""

    @pytest.mark.asyncio
    async def test_full_sync_waits_for_running_delta(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, graph, _ = _make_engine(hass, [simple_person])
        gate = _first_call_gate(google.fetch_contacts)
        hass.async_add_executor_job = AsyncMock(side_effect=gate)

        first = asyncio.create_task(engine.async_sync())
        await gate.reached.wait()  # delta run is suspended inside its fetch

        second = asyncio.create_task(engine.async_full_sync())
        await _settle()

        # Its own fetch would not block, so reaching the Outlook snapshot here
        # would mean the full sync overtook the running delta.
        graph.list_contacts_raw.assert_not_called()
        assert not second.done()

        gate.release.set()
        await asyncio.wait_for(asyncio.gather(first, second), timeout=5)
        graph.list_contacts_raw.assert_called_once()

    @pytest.mark.asyncio
    async def test_dry_run_waits_for_running_delta(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, _, _ = _make_engine(hass, [simple_person])
        gate = _first_call_gate(google.fetch_contacts)
        hass.async_add_executor_job = AsyncMock(side_effect=gate)

        first = asyncio.create_task(engine.async_sync())
        await gate.reached.wait()

        plan = asyncio.create_task(engine.async_plan())
        await _settle()

        # Unlocked, the dry run would sail through to completion here.
        assert not plan.done()

        gate.release.set()
        await asyncio.wait_for(asyncio.gather(first, plan), timeout=5)

    @pytest.mark.asyncio
    async def test_full_sync_and_dry_run_are_serialised(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, _, _ = _make_engine(hass, [simple_person])
        gate = _first_call_gate(google.fetch_contacts)
        hass.async_add_executor_job = AsyncMock(side_effect=gate)

        first = asyncio.create_task(engine.async_full_sync())
        await gate.reached.wait()

        plan = asyncio.create_task(engine.async_plan())
        await _settle()
        assert not plan.done()

        gate.release.set()
        await asyncio.wait_for(asyncio.gather(first, plan), timeout=5)

    @pytest.mark.asyncio
    async def test_lock_is_released_after_an_exception(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, _, _ = _make_engine(hass, [simple_person])
        google.fetch_contacts.side_effect = RuntimeError("boom")

        with pytest.raises(RuntimeError):
            await engine.async_sync()

        # A later caller must still be able to acquire the lock.
        google.fetch_contacts.side_effect = None
        google.fetch_contacts.return_value = ([simple_person], "token_new")
        result = await engine.async_sync()
        assert result.created == 1

    @pytest.mark.asyncio
    async def test_no_double_create_from_overlapping_runs(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """Overlapping runs must create an unmapped contact at most once.

        Unlocked, the second run decides CREATE from a store snapshot taken
        before the first run saved its mapping, and Outlook ends up with two
        copies of the same person.
        """
        engine, _, graph, _ = _make_engine(hass, [simple_person])
        gate = _first_call_gate(graph.create_contact)
        hass.async_add_executor_job = AsyncMock(side_effect=gate)

        first = asyncio.create_task(engine.async_sync())
        await gate.reached.wait()  # first run is mid-create, mapping not saved

        second = asyncio.create_task(engine.async_sync())
        await _settle()

        gate.release.set()
        await asyncio.wait_for(asyncio.gather(first, second), timeout=5)

        # Having acquired the lock, the second run re-reads the store, finds the
        # mapping and hash the first run wrote, and skips instead of creating.
        assert graph.create_contact.call_count == 1
        second_result = second.result()
        assert second_result.created == 0
        assert second_result.skipped == 1


class TestCursorSemantics:
    """#33: no valid cursor always means full reconciliation."""

    @pytest.mark.asyncio
    async def test_sync_without_token_uses_full_pipeline(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, graph, _ = _make_engine(
            hass, [simple_person], stored_sync_token=None
        )
        await engine.async_sync()

        # The Outlook snapshot is only loaded by the full pipeline.
        graph.list_contacts_raw.assert_called_once()
        google.fetch_contacts.assert_called_once_with(None)

    @pytest.mark.asyncio
    async def test_expired_token_uses_full_pipeline(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, google, graph, _ = _make_engine(hass, [simple_person])
        google.fetch_contacts.side_effect = [
            SyncTokenExpiredError("expired"),
            ([simple_person], "token_fresh"),
        ]

        await engine.async_sync()

        graph.list_contacts_raw.assert_called_once()
        assert google.fetch_contacts.call_args_list[-1][0][0] is None

    @pytest.mark.asyncio
    async def test_successful_full_sync_stores_fresh_token(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, _, store = _make_engine(
            hass, [simple_person], sync_token="tok_fresh"
        )
        await engine.async_full_sync()
        assert store.sync_token == "tok_fresh"

    @pytest.mark.asyncio
    async def test_failed_full_sync_clears_the_cursor(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        """A stale delta cursor would never re-deliver the failed contact."""
        engine, _, graph, store = _make_engine(
            hass, [simple_person], sync_token="tok_fresh"
        )
        graph.create_contact.side_effect = RuntimeError("boom")

        result = await engine.async_full_sync()

        assert result.failed == 1
        assert store.sync_token is None
        store.async_save.assert_called_once()

    @pytest.mark.asyncio
    async def test_run_after_failed_full_sync_is_full_again(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass, [simple_person], sync_token="tok_fresh"
        )
        graph.create_contact.side_effect = RuntimeError("boom")
        await engine.async_full_sync()
        assert store.sync_token is None

        graph.create_contact.side_effect = None
        graph.create_contact.return_value = "graph-id-001"
        graph.list_contacts_raw.reset_mock()

        await engine.async_sync()
        graph.list_contacts_raw.assert_called_once()

    @pytest.mark.asyncio
    async def test_failed_delta_keeps_its_cursor(
        self, hass: MagicMock, simple_person: dict[str, Any]
    ) -> None:
        engine, _, graph, store = _make_engine(
            hass, [simple_person], sync_token="tok_new"
        )
        graph.create_contact.side_effect = RuntimeError("boom")

        await engine.async_sync()
        assert store.sync_token == "token_stored"


class TestMicrosoftTokenCheck:
    """#38: every run checks the Microsoft sign-in before doing anything."""

    @pytest.mark.asyncio
    @pytest.mark.parametrize("run", ["async_sync", "async_full_sync", "async_plan"])
    async def test_every_run_checks_first(self, hass: MagicMock, run: str) -> None:
        engine, google, graph, _ = _make_engine(hass, [])
        order: list[str] = []
        graph.ensure_token.side_effect = lambda: order.append("ensure")
        google.fetch_contacts.side_effect = lambda *_: (
            order.append("fetch") or ([], "tok")
        )

        await getattr(engine, run)()

        assert order[:2] == ["ensure", "fetch"]

    @pytest.mark.asyncio
    @pytest.mark.parametrize("run", ["async_sync", "async_full_sync", "async_plan"])
    async def test_rejected_sign_in_stops_the_run(
        self, hass: MagicMock, run: str
    ) -> None:
        engine, google, graph, _ = _make_engine(hass, [])
        graph.ensure_token.side_effect = MSAuthError("invalid_client")
        persisted: list[bool] = []

        async def _persist() -> None:
            persisted.append(True)

        engine._persist_tokens_cb = _persist

        with pytest.raises(MSAuthError):
            await getattr(engine, run)()

        google.fetch_contacts.assert_not_called()
        assert persisted == [True]

    @pytest.mark.asyncio
    async def test_expired_secret_surfaces_in_a_run_without_changes(
        self, hass: MagicMock
    ) -> None:
        """The real client: no contact changed, so no Graph request is due."""
        engine, google, _, _ = _make_engine(hass, [])
        con = MagicMock()
        con.token_backend.get_access_token.return_value = {"expires_on": "0"}
        engine._graph = GraphContactsClient(MagicMock(con=con))

        with (
            patch.object(
                graph_client_module,
                "refresh_access_token",
                return_value=(CredentialStatus.INVALID_CLIENT, None),
            ),
            pytest.raises(MSAuthError),
        ):
            await engine.async_sync()

        con.get.assert_not_called()
        con.post.assert_not_called()


def _mapped_people(count: int) -> tuple[list[dict[str, Any]], dict[str, str]]:
    """*count* Google contacts, each already mapped to an Outlook contact."""
    people = [
        {
            "resourceName": f"people/c{i}",
            "names": [{"displayName": f"Person {i}"}],
        }
        for i in range(count)
    ]
    mapping = {f"people/c{i}": f"graph-{i}" for i in range(count)}
    return people, mapping


class TestSystematicFailure:
    """Every write failing must not turn into an endless loop of full resyncs."""

    @pytest.mark.asyncio
    async def test_full_sync_stops_when_every_write_fails(
        self, hass: MagicMock
    ) -> None:
        people, mapping = _mapped_people(ABORT_AFTER_FAILURES + 5)
        outlook = [{"id": cid} for cid in mapping.values()] + [{"id": "graph-orphan"}]
        mapping["people/gone"] = "graph-orphan"
        engine, _, graph, store = _make_engine(
            hass,
            people,
            sync_token="tok_fresh",
            existing_mapping=mapping,
            outlook_raw=outlook,
            delete_removed=True,
        )
        graph.update_contact.side_effect = RuntimeError("HTTP 500")

        result = await engine.async_full_sync()

        assert result.aborted is True
        assert graph.update_contact.call_count == ABORT_AFTER_FAILURES
        assert result.failed == ABORT_AFTER_FAILURES
        assert "aborted" in result.errors[-1]
        # Discarded like any failed full resync, never replaced by the new one.
        assert store.sync_token is None
        # No orphan delete on top of a systematic failure.
        graph.delete_contact.assert_not_called()
        store.async_save.assert_called_once()

    @pytest.mark.asyncio
    async def test_run_after_an_aborted_full_sync_retries_every_contact(
        self, hass: MagicMock
    ) -> None:
        """Keeping the old cursor made the next delta skip the unreached contacts."""
        count = ABORT_AFTER_FAILURES + 5
        people, mapping = _mapped_people(count)
        engine, google, graph, _ = _make_engine(
            hass,
            people,
            sync_token="tok_fresh",
            existing_mapping=mapping,
            outlook_raw=[{"id": cid} for cid in mapping.values()],
        )
        graph.update_contact.side_effect = RuntimeError("HTTP 500")
        await engine.async_full_sync()

        graph.update_contact.reset_mock(side_effect=True)
        google.fetch_contacts.reset_mock()

        result = await engine.async_sync()

        google.fetch_contacts.assert_called_once_with(None)
        assert graph.update_contact.call_count == count
        assert result.updated == count
        assert result.failed == 0

    @pytest.mark.asyncio
    async def test_one_success_means_failures_are_per_contact(
        self, hass: MagicMock
    ) -> None:
        people, mapping = _mapped_people(ABORT_AFTER_FAILURES + 5)
        engine, _, graph, store = _make_engine(
            hass,
            people,
            sync_token="tok_fresh",
            existing_mapping=mapping,
            outlook_raw=[{"id": cid} for cid in mapping.values()],
        )
        graph.update_contact.side_effect = [None] + [RuntimeError("boom")] * (
            ABORT_AFTER_FAILURES + 4
        )

        result = await engine.async_full_sync()

        assert result.aborted is False
        assert graph.update_contact.call_count == ABORT_AFTER_FAILURES + 5
        assert result.updated == 1
        assert result.failed == ABORT_AFTER_FAILURES + 4
        assert store.sync_token is None

    @pytest.mark.asyncio
    async def test_reported_errors_are_capped(self, hass: MagicMock) -> None:
        count = MAX_REPORTED_ERRORS + 10
        people, mapping = _mapped_people(count)
        engine, _, graph, _ = _make_engine(
            hass,
            people,
            existing_mapping=mapping,
            outlook_raw=[{"id": cid} for cid in mapping.values()],
        )
        graph.update_contact.side_effect = [None] + [RuntimeError("boom")] * (count - 1)

        result = await engine.async_full_sync()

        assert result.failed == count - 1
        assert len(result.errors) == MAX_REPORTED_ERRORS

    @pytest.mark.asyncio
    async def test_delta_sync_stops_and_keeps_its_cursor(self, hass: MagicMock) -> None:
        people, mapping = _mapped_people(ABORT_AFTER_FAILURES + 5)
        engine, _, graph, store = _make_engine(
            hass, people, sync_token="tok_new", existing_mapping=mapping
        )
        graph.update_contact.side_effect = RuntimeError("HTTP 500")

        result = await engine.async_sync()

        assert result.aborted is True
        assert graph.update_contact.call_count == ABORT_AFTER_FAILURES
        assert store.sync_token == "token_stored"
