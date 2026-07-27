"""Tests for the sensor, button and diagnostics platforms.

Entities are built with __new__ and given a mocked coordinator: instantiating
them normally drags in HA's entity platform machinery, which needs a real hass.
What matters here is how each entity reads the coordinator result.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from unittest.mock import AsyncMock, MagicMock

import pytest

from custom_components.google_outlook_contacts_sync.button import (
    DryRunButton,
    FullResyncButton,
    SyncNowButton,
)
from custom_components.google_outlook_contacts_sync.const import (
    DATA_COORDINATOR,
    DATA_STORE,
    DOMAIN,
)
from custom_components.google_outlook_contacts_sync.diagnostics import (
    async_get_config_entry_diagnostics,
)
from custom_components.google_outlook_contacts_sync.sensor import (
    DryRunPreviewSensor,
    DuplicatesRemovedSensor,
    LastSyncResultSensor,
    LastSyncSensor,
    NextSyncSensor,
    SyncedContactsSensor,
)
from custom_components.google_outlook_contacts_sync.sync.engine import (
    PlannedChange,
    SyncPlan,
    SyncResult,
)

FINISHED = datetime(2026, 7, 27, 10, 0, tzinfo=UTC)


def _entity(cls, data, **attrs):
    """Build an entity without running the HA entity constructor."""
    coordinator = MagicMock()
    coordinator.data = data
    coordinator.update_interval = attrs.pop("update_interval", timedelta(hours=24))
    coordinator.last_update_success = attrs.pop("last_update_success", True)
    coordinator.last_exception = None

    entity = cls.__new__(cls)
    entity.coordinator = coordinator
    for key, value in attrs.items():
        setattr(entity, key, value)
    return entity


class TestLastSyncSensor:
    def test_reports_the_finish_time(self):
        result = SyncResult(finished_at=FINISHED)
        assert _entity(LastSyncSensor, result).native_value == FINISHED

    def test_none_before_the_first_run(self):
        assert _entity(LastSyncSensor, None).native_value is None


class TestNextSyncSensor:
    def test_adds_the_update_interval(self):
        result = SyncResult(finished_at=FINISHED)
        sensor = _entity(NextSyncSensor, result, update_interval=timedelta(hours=6))
        assert sensor.native_value == FINISHED + timedelta(hours=6)

    def test_none_without_a_finish_time(self):
        assert _entity(NextSyncSensor, SyncResult()).native_value is None

    def test_none_without_an_interval(self):
        result = SyncResult(finished_at=FINISHED)
        sensor = _entity(NextSyncSensor, result, update_interval=None)
        assert sensor.native_value is None


class TestLastSyncResultSensor:
    def test_summary_and_attributes(self):
        result = SyncResult(created=2, updated=1, failed=1, errors=["people/x: boom"])
        sensor = _entity(LastSyncResultSensor, result)

        assert "2 created" in sensor.native_value
        assert sensor.extra_state_attributes["failed"] == 1
        assert sensor.extra_state_attributes["errors"] == ["people/x: boom"]

    def test_empty_attributes_before_the_first_run(self):
        assert _entity(LastSyncResultSensor, None).extra_state_attributes == {}


class TestSyncedContactsSensor:
    def test_counts_tracked_resource_names(self):
        store = MagicMock()
        store.all_resource_names.return_value = {"people/a", "people/b"}
        sensor = _entity(SyncedContactsSensor, SyncResult(), _store=store)
        assert sensor.native_value == 2

    def test_drops_after_contacts_are_untracked(self):
        """#18: untracked contacts must stop being counted."""
        store = MagicMock()
        store.all_resource_names.return_value = set()
        sensor = _entity(SyncedContactsSensor, SyncResult(), _store=store)
        assert sensor.native_value == 0


class TestDuplicatesRemovedSensor:
    def test_reports_the_count_and_groups(self):
        result = SyncResult(duplicates_removed=2, duplicates=["bob jones"])
        sensor = _entity(DuplicatesRemovedSensor, result)

        assert sensor.native_value == 2
        assert sensor.extra_state_attributes == {"groups": ["bob jones"]}

    def test_no_group_attribute_when_nothing_was_removed(self):
        sensor = _entity(DuplicatesRemovedSensor, SyncResult())
        assert sensor.extra_state_attributes == {}

    def test_none_before_the_first_run(self):
        assert _entity(DuplicatesRemovedSensor, None).native_value is None


class TestDryRunPreviewSensor:
    def test_reports_the_planned_total(self):
        plan = SyncPlan(
            to_create=[PlannedChange("people/a", "A")],
            to_update=[PlannedChange("people/b", "B")],
        )
        assert _entity(DryRunPreviewSensor, plan).native_value == 2

    def test_none_before_a_dry_run(self):
        assert _entity(DryRunPreviewSensor, None).native_value is None

    def test_empty_attributes_before_a_dry_run(self):
        assert _entity(DryRunPreviewSensor, None).extra_state_attributes == {}


class TestButtons:
    """#17: buttons and services must use the same refresh semantics."""

    @staticmethod
    def _button(cls):
        coordinator = MagicMock()
        coordinator.async_request_refresh = AsyncMock()
        coordinator.async_full_resync = AsyncMock()
        button = cls.__new__(cls)
        button._coordinator = coordinator
        return button, coordinator

    @pytest.mark.asyncio
    async def test_sync_now_uses_the_debounced_refresh(self):
        button, coordinator = self._button(SyncNowButton)
        await button.async_press()
        coordinator.async_request_refresh.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_full_resync_runs_a_full_reconciliation(self):
        button, coordinator = self._button(FullResyncButton)
        await button.async_press()
        coordinator.async_full_resync.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_dry_run_refreshes_its_own_coordinator(self):
        button, coordinator = self._button(DryRunButton)
        await button.async_press()
        coordinator.async_request_refresh.assert_awaited_once()


class TestDiagnostics:
    @pytest.mark.asyncio
    async def test_credentials_are_redacted(self):
        coordinator = MagicMock()
        coordinator.data = SyncResult(created=1, finished_at=FINISHED)
        coordinator.last_update_success = True
        coordinator.last_exception = None
        coordinator.update_interval = timedelta(hours=24)

        store = MagicMock()
        store.all_resource_names.return_value = {"people/a"}
        store.sync_token = "tok"

        hass = MagicMock()
        entry = MagicMock()
        entry.entry_id = "entry-1"
        # Distinct sentinels so a hit cannot come from a key name like
        # "ms_client_secret".
        entry.data = {
            "ms_client_id": "SENTINEL-MS-ID",
            "ms_client_secret": "SENTINEL-MS-SECRET",
            "google_client_secret": "SENTINEL-GOOGLE-SECRET",
            "ms_token": {"access_token": "SENTINEL-MS-TOKEN"},
            "google_token": {"refresh_token": "SENTINEL-GOOGLE-REFRESH"},
        }
        entry.options = {"sync_interval_hours": 12}
        hass.data = {
            DOMAIN: {
                "entry-1": {DATA_COORDINATOR: coordinator, DATA_STORE: store},
            }
        }

        result = await async_get_config_entry_diagnostics(hass, entry)

        dumped = str(result)
        assert "SENTINEL" not in dumped, "a credential survived redaction"

        assert result["entry"]["options"] == {"sync_interval_hours": 12}
        assert result["store"] == {"synced_contacts": 1, "has_sync_token": True}
        assert result["coordinator"]["last_result"]["created"] == 1
