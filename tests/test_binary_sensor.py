"""Tests for binary_sensor.py — the sync problem indicator."""

from __future__ import annotations

from unittest.mock import MagicMock

from custom_components.google_outlook_contacts_sync.binary_sensor import (
    SyncProblemBinarySensor,
)
from custom_components.google_outlook_contacts_sync.sync.engine import SyncResult


def _sensor(data: SyncResult | None, last_update_success: bool = True):
    coordinator = MagicMock()
    coordinator.data = data
    coordinator.last_update_success = last_update_success
    coordinator.last_exception = None
    entry = MagicMock()
    entry.entry_id = "entry-1"
    sensor = SyncProblemBinarySensor.__new__(SyncProblemBinarySensor)
    sensor.coordinator = coordinator
    return sensor


class TestSyncProblem:
    def test_off_on_a_clean_run(self):
        assert _sensor(SyncResult()).is_on is False

    def test_on_when_the_coordinator_update_failed(self):
        assert _sensor(SyncResult(), last_update_success=False).is_on is True

    def test_on_for_a_partial_dedup_failure(self):
        """#27: a failed dedup op must light this up, not just fill errors."""
        result = SyncResult(failed=1, errors=["Dedup listing failed: boom"])
        assert _sensor(result).is_on is True

    def test_errors_are_exposed_as_attributes(self):
        result = SyncResult(failed=1, errors=["Dedup delete failed: boom"])
        assert _sensor(result).extra_state_attributes["errors"] == [
            "Dedup delete failed: boom"
        ]
