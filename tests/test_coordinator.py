"""Tests for coordinator.py — error routing between the engine and HA.

Coordinators are built with __new__: DataUpdateCoordinator.__init__ reaches for
HA's frame helper, which needs a real hass. The logic under test is which
exception type maps to which HA outcome.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from homeassistant.exceptions import ConfigEntryAuthFailed
from homeassistant.helpers.update_coordinator import (
    DataUpdateCoordinator,
    UpdateFailed,
)

from custom_components.google_outlook_contacts_sync.const import (
    DATA_REAUTH_PROVIDER,
    DOMAIN,
)
from custom_components.google_outlook_contacts_sync.coordinator import (
    ContactSyncCoordinator,
    DryRunCoordinator,
)
from custom_components.google_outlook_contacts_sync.sync.engine import (
    SyncPlan,
    SyncResult,
)
from custom_components.google_outlook_contacts_sync.sync.google_client import (
    GoogleAuthError,
)
from custom_components.google_outlook_contacts_sync.sync.graph_client import MSAuthError


def _coordinator(engine: MagicMock) -> ContactSyncCoordinator:
    hass = MagicMock()
    hass.data = {}
    entry = MagicMock()
    entry.entry_id = "entry-1"

    coordinator = ContactSyncCoordinator.__new__(ContactSyncCoordinator)
    coordinator.hass = hass
    coordinator._entry = entry
    coordinator._engine = engine
    coordinator.async_set_updated_data = MagicMock()
    coordinator.async_set_update_error = MagicMock()
    return coordinator


def _dry_run(engine: MagicMock) -> DryRunCoordinator:
    coordinator = DryRunCoordinator.__new__(DryRunCoordinator)
    coordinator.hass = MagicMock()
    coordinator.hass.data = {}
    coordinator._entry = MagicMock()
    coordinator._entry.entry_id = "entry-1"
    coordinator._engine = engine
    return coordinator


class TestDeltaUpdate:
    @pytest.mark.asyncio
    async def test_returns_the_engine_result(self):
        engine = MagicMock()
        engine.async_sync = AsyncMock(return_value=SyncResult(created=3))

        result = await _coordinator(engine)._async_update_data()

        assert result.created == 3

    @pytest.mark.asyncio
    async def test_google_auth_error_becomes_config_entry_auth_failed(self):
        engine = MagicMock()
        engine.async_sync = AsyncMock(side_effect=GoogleAuthError("invalid_grant"))
        coordinator = _coordinator(engine)

        with pytest.raises(ConfigEntryAuthFailed):
            await coordinator._async_update_data()

        provider = coordinator.hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"]
        assert provider == "google"

    @pytest.mark.asyncio
    async def test_microsoft_auth_error_routes_to_the_ms_provider(self):
        engine = MagicMock()
        engine.async_sync = AsyncMock(side_effect=MSAuthError("token expired"))
        coordinator = _coordinator(engine)

        with pytest.raises(ConfigEntryAuthFailed):
            await coordinator._async_update_data()

        provider = coordinator.hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"]
        assert provider == "ms"

    @pytest.mark.asyncio
    async def test_other_errors_become_update_failed(self):
        """#14: a failed fetch must surface as a failed update, not a clean run."""
        engine = MagicMock()
        engine.async_sync = AsyncMock(side_effect=ConnectionError("network down"))

        with pytest.raises(UpdateFailed):
            await _coordinator(engine)._async_update_data()


class TestFullResync:
    @pytest.mark.asyncio
    async def test_success_publishes_the_result(self):
        engine = MagicMock()
        engine.async_full_sync = AsyncMock(return_value=SyncResult(updated=2))
        coordinator = _coordinator(engine)

        await coordinator.async_full_resync()

        published = coordinator.async_set_updated_data.call_args.args[0]
        assert published.updated == 2

    @pytest.mark.asyncio
    async def test_auth_error_sets_the_error_and_starts_reauth(self):
        engine = MagicMock()
        engine.async_full_sync = AsyncMock(side_effect=MSAuthError("token expired"))
        coordinator = _coordinator(engine)

        await coordinator.async_full_resync()

        coordinator.async_set_update_error.assert_called_once()
        coordinator.async_set_updated_data.assert_not_called()
        coordinator._entry.async_start_reauth.assert_called_once()

    @pytest.mark.asyncio
    async def test_other_errors_set_the_update_error(self):
        engine = MagicMock()
        engine.async_full_sync = AsyncMock(side_effect=RuntimeError("boom"))
        coordinator = _coordinator(engine)

        await coordinator.async_full_resync()

        coordinator.async_set_update_error.assert_called_once()
        coordinator.async_set_updated_data.assert_not_called()


class TestDryRunCoordinator:
    @pytest.mark.asyncio
    async def test_returns_the_plan(self):
        engine = MagicMock()
        engine.async_plan = AsyncMock(return_value=SyncPlan())

        assert isinstance(await _dry_run(engine)._async_update_data(), SyncPlan)

    @pytest.mark.asyncio
    async def test_errors_become_update_failed(self):
        engine = MagicMock()
        engine.async_plan = AsyncMock(side_effect=RuntimeError("boom"))

        with pytest.raises(UpdateFailed):
            await _dry_run(engine)._async_update_data()

    @pytest.mark.asyncio
    @pytest.mark.parametrize(
        ("exc", "provider"),
        [(MSAuthError("invalid_client"), "ms"), (GoogleAuthError("gone"), "google")],
    )
    async def test_auth_errors_start_reauth(self, exc, provider):
        """#38: a dry run checks the Microsoft sign-in too."""
        engine = MagicMock()
        engine.async_plan = AsyncMock(side_effect=exc)
        coordinator = _dry_run(engine)

        with pytest.raises(ConfigEntryAuthFailed):
            await coordinator._async_update_data()

        stored = coordinator.hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"]
        assert stored == provider

    def test_reauth_wiring_reaches_home_assistant(self):
        """HA starts reauth from ConfigEntryAuthFailed only with config_entry."""
        entry = MagicMock()
        entry.entry_id = "entry-1"
        with patch.object(DataUpdateCoordinator, "__init__", return_value=None) as init:
            DryRunCoordinator(MagicMock(), entry, MagicMock())

        assert init.call_args.kwargs["config_entry"] is entry
