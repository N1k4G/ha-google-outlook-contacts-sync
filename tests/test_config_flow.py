"""Tests for config_flow.py."""

from __future__ import annotations

import pytest
import voluptuous as vol

from custom_components.google_outlook_contacts_sync.config_flow import _OPTIONS_SCHEMA
from custom_components.google_outlook_contacts_sync.const import (
    CONF_DELETE_REMOVED,
    CONF_SYNC_INTERVAL_HOURS,
    DEFAULT_DELETE_REMOVED,
    DEFAULT_SYNC_INTERVAL_HOURS,
)


class TestOptionsFlow:
    def test_options_schema_default_values(self) -> None:
        result = _OPTIONS_SCHEMA(
            {
                CONF_SYNC_INTERVAL_HOURS: DEFAULT_SYNC_INTERVAL_HOURS,
                CONF_DELETE_REMOVED: DEFAULT_DELETE_REMOVED,
            }
        )
        assert result[CONF_SYNC_INTERVAL_HOURS] == DEFAULT_SYNC_INTERVAL_HOURS
        assert result[CONF_DELETE_REMOVED] == DEFAULT_DELETE_REMOVED

    def test_options_schema_rejects_zero_interval(self) -> None:
        with pytest.raises(vol.Invalid):
            _OPTIONS_SCHEMA(
                {CONF_SYNC_INTERVAL_HOURS: 0, CONF_DELETE_REMOVED: False}
            )

    def test_options_schema_rejects_interval_above_max(self) -> None:
        with pytest.raises(vol.Invalid):
            _OPTIONS_SCHEMA(
                {CONF_SYNC_INTERVAL_HOURS: 200, CONF_DELETE_REMOVED: False}
            )
