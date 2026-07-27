"""Tests for __init__.py — config entry lifecycle.

The Home Assistant boundary is mocked rather than driven through
pytest-homeassistant-custom-component: HA's runner imports fcntl, so that
plugin cannot load on Windows, and these paths are about our own ordering and
bookkeeping rather than HA internals.
"""

from __future__ import annotations

from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from homeassistant.exceptions import ConfigEntryAuthFailed

from custom_components.google_outlook_contacts_sync import (
    _async_entry_updated,
    _async_persist_tokens,
    async_setup_entry,
    async_unload_entry,
)
from custom_components.google_outlook_contacts_sync.const import (
    DATA_COORDINATOR,
    DATA_OPTIONS_SNAPSHOT,
    DATA_REAUTH_PROVIDER,
    DOMAIN,
    SERVICE_FULL_SYNC,
    SERVICE_SYNC_NOW,
)

GOOGLE_TOKEN = {
    "client_id": "gid",
    "client_secret": "gsecret",
    "refresh_token": "grefresh",
    "token": "gaccess",
}


def _make_hass() -> MagicMock:
    hass = MagicMock()
    hass.data = {}
    hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *a: fn(*a))
    hass.config_entries.async_forward_entry_setups = AsyncMock()
    hass.config_entries.async_reload = AsyncMock()
    hass.config_entries.async_update_entry = MagicMock()
    hass.services.has_service = MagicMock(return_value=False)
    hass.services.async_register = MagicMock()
    hass.services.async_remove = MagicMock()
    return hass


def _make_entry(
    google_token: dict[str, Any] | None = None,
    options: dict[str, Any] | None = None,
) -> MagicMock:
    entry = MagicMock()
    entry.entry_id = "entry-1"
    entry.data = {
        "ms_client_id": "mid",
        "ms_client_secret": "msecret",
        "ms_token": {"access_token": "mtok"},
        "google_token": GOOGLE_TOKEN if google_token is None else google_token,
    }
    entry.options = options or {}

    created: list[Any] = []

    def _create_background_task(_hass, coro, _name):
        created.append(coro)
        coro.close()  # never actually run in these tests

    entry.async_create_background_task = MagicMock(side_effect=_create_background_task)
    entry.background_tasks = created
    return entry


def _dummy_coro():
    """A fresh awaitable for the mocked coordinator entry points."""

    async def _noop() -> None:
        return None

    return _noop()


def _make_coordinator() -> MagicMock:
    coordinator = MagicMock()
    coordinator.async_full_resync = MagicMock(side_effect=lambda: _dummy_coro())
    coordinator.async_request_refresh = MagicMock(side_effect=lambda: _dummy_coro())
    return coordinator


async def _setup(hass, entry, authenticated: bool = True, sync_token=None):
    """Run async_setup_entry with the external boundary mocked.

    The coordinators are mocked too: DataUpdateCoordinator.__init__ reaches for
    HA's frame helper, which needs a real hass. Their behaviour is covered by
    the engine and coordinator tests; what matters here is setup ordering.
    """
    account = MagicMock()
    account.is_authenticated = authenticated
    backend = MagicMock()
    backend.get_token_dict.return_value = {"access_token": "mtok"}

    store = MagicMock()
    store.async_load = AsyncMock()
    store.sync_token = sync_token

    coordinator = _make_coordinator()

    with (
        patch(
            "custom_components.google_outlook_contacts_sync.build_account",
            return_value=(account, backend),
        ),
        patch(
            "custom_components.google_outlook_contacts_sync."
            "GoogleContactsClient.from_token_dict",
            return_value=MagicMock(),
        ),
        patch(
            "custom_components.google_outlook_contacts_sync.ContactSyncStore",
            return_value=store,
        ),
        patch(
            "custom_components.google_outlook_contacts_sync.ContactSyncCoordinator",
            return_value=coordinator,
        ),
        patch(
            "custom_components.google_outlook_contacts_sync.DryRunCoordinator",
            return_value=MagicMock(),
        ),
    ):
        result = await async_setup_entry(hass, entry)
    return result, coordinator


class TestGoogleTokenValidation:
    """#16: a missing Google token must route to reauth, not crash setup."""

    @pytest.mark.asyncio
    async def test_empty_google_token_raises_auth_failed(self):
        hass, entry = _make_hass(), _make_entry(google_token={})

        with pytest.raises(ConfigEntryAuthFailed):
            await _setup(hass, entry)

        assert hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"] == "google"

    @pytest.mark.asyncio
    async def test_incomplete_google_token_raises_auth_failed(self):
        incomplete = {"client_id": "gid", "client_secret": "gsecret"}
        hass, entry = _make_hass(), _make_entry(google_token=incomplete)

        with pytest.raises(ConfigEntryAuthFailed):
            await _setup(hass, entry)

        assert hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"] == "google"

    @pytest.mark.asyncio
    async def test_unauthenticated_microsoft_routes_to_ms_reauth(self):
        hass, entry = _make_hass(), _make_entry()

        with pytest.raises(ConfigEntryAuthFailed):
            await _setup(hass, entry, authenticated=False)

        assert hass.data[DOMAIN][DATA_REAUTH_PROVIDER]["entry-1"] == "ms"


class TestInitialSync:
    """#19: setup publishes entities first and syncs in the background."""

    @pytest.mark.asyncio
    async def test_setup_does_not_await_the_sync(self):
        hass, entry = _make_hass(), _make_entry()

        ok, _ = await _setup(hass, entry)
        assert ok is True

        # Runtime data and platforms are ready before the sync is scheduled.
        assert DATA_COORDINATOR in hass.data[DOMAIN]["entry-1"]
        hass.config_entries.async_forward_entry_setups.assert_awaited_once()
        entry.async_create_background_task.assert_called_once()

    @pytest.mark.asyncio
    async def test_without_cursor_runs_a_full_resync(self):
        hass, entry = _make_hass(), _make_entry()

        _, coordinator = await _setup(hass, entry, sync_token=None)

        coordinator.async_full_resync.assert_called_once()
        coordinator.async_request_refresh.assert_not_called()

    @pytest.mark.asyncio
    async def test_with_cursor_runs_a_delta_refresh(self):
        hass, entry = _make_hass(), _make_entry()

        _, coordinator = await _setup(hass, entry, sync_token="existing-token")

        coordinator.async_request_refresh.assert_called_once()
        coordinator.async_full_resync.assert_not_called()


class TestServices:
    """#17: services are registered once and cleaned up on unload."""

    @pytest.mark.asyncio
    async def test_services_registered_on_setup(self):
        hass, entry = _make_hass(), _make_entry()
        await _setup(hass, entry)

        registered = {c.args[1] for c in hass.services.async_register.call_args_list}
        assert registered == {SERVICE_SYNC_NOW, SERVICE_FULL_SYNC}

    @pytest.mark.asyncio
    async def test_services_not_registered_twice(self):
        hass, entry = _make_hass(), _make_entry()
        hass.services.has_service.return_value = True

        await _setup(hass, entry)

        hass.services.async_register.assert_not_called()

    @pytest.mark.asyncio
    async def test_services_removed_when_last_entry_unloads(self):
        hass, entry = _make_hass(), _make_entry()
        await _setup(hass, entry)

        hass.config_entries.async_unload_platforms = AsyncMock(return_value=True)
        assert await async_unload_entry(hass, entry) is True

        removed = {c.args[1] for c in hass.services.async_remove.call_args_list}
        assert removed == {SERVICE_SYNC_NOW, SERVICE_FULL_SYNC}
        assert "entry-1" not in hass.data[DOMAIN]

    @pytest.mark.asyncio
    async def test_services_kept_while_another_entry_remains(self):
        hass, entry = _make_hass(), _make_entry()
        await _setup(hass, entry)
        hass.data[DOMAIN]["entry-2"] = {DATA_COORDINATOR: MagicMock()}

        hass.config_entries.async_unload_platforms = AsyncMock(return_value=True)
        await async_unload_entry(hass, entry)

        hass.services.async_remove.assert_not_called()


class TestTokenPersistence:
    """#7: refreshed tokens survive a restart, without a reload loop."""

    @staticmethod
    def _parts(ms_token=None, google_token=None):
        hass, entry = _make_hass(), _make_entry()
        backend = MagicMock()
        backend.get_token_dict.return_value = (
            ms_token if ms_token is not None else entry.data["ms_token"]
        )
        client = MagicMock()
        client.get_token_dict.return_value = (
            google_token if google_token is not None else entry.data["google_token"]
        )
        return hass, entry, backend, client

    @pytest.mark.asyncio
    async def test_no_change_writes_nothing(self):
        hass, entry, backend, client = self._parts()
        await _async_persist_tokens(hass, entry, backend, client)
        hass.config_entries.async_update_entry.assert_not_called()

    @pytest.mark.asyncio
    async def test_microsoft_change_is_written(self):
        hass, entry, backend, client = self._parts(ms_token={"access_token": "new"})
        await _async_persist_tokens(hass, entry, backend, client)

        hass.config_entries.async_update_entry.assert_called_once()
        data = hass.config_entries.async_update_entry.call_args.kwargs["data"]
        assert data["ms_token"] == {"access_token": "new"}
        assert data["ms_client_id"] == "mid"  # untouched

    @pytest.mark.asyncio
    async def test_google_change_is_written(self):
        new_google = {**GOOGLE_TOKEN, "token": "refreshed"}
        hass, entry, backend, client = self._parts(google_token=new_google)
        await _async_persist_tokens(hass, entry, backend, client)

        data = hass.config_entries.async_update_entry.call_args.kwargs["data"]
        assert data["google_token"]["token"] == "refreshed"

    @pytest.mark.asyncio
    async def test_both_changed_write_once(self):
        hass, entry, backend, client = self._parts(
            ms_token={"access_token": "new"},
            google_token={**GOOGLE_TOKEN, "token": "refreshed"},
        )
        await _async_persist_tokens(hass, entry, backend, client)

        assert hass.config_entries.async_update_entry.call_count == 1


class TestEntryUpdateListener:
    """#7: token writes must not reload; option changes must."""

    @pytest.mark.asyncio
    async def test_unchanged_options_do_not_reload(self):
        hass, entry = _make_hass(), _make_entry(options={"sync_interval_hours": 12})
        hass.data[DOMAIN] = {
            "entry-1": {
                DATA_COORDINATOR: MagicMock(),
                DATA_OPTIONS_SNAPSHOT: {"sync_interval_hours": 12},
            }
        }

        await _async_entry_updated(hass, entry)

        hass.config_entries.async_reload.assert_not_called()

    @pytest.mark.asyncio
    async def test_changed_options_reload_once(self):
        hass, entry = _make_hass(), _make_entry(options={"sync_interval_hours": 6})
        hass.data[DOMAIN] = {
            "entry-1": {
                DATA_COORDINATOR: MagicMock(),
                DATA_OPTIONS_SNAPSHOT: {"sync_interval_hours": 12},
            }
        }

        await _async_entry_updated(hass, entry)

        hass.config_entries.async_reload.assert_awaited_once_with("entry-1")

    @pytest.mark.asyncio
    async def test_unknown_entry_is_ignored(self):
        hass, entry = _make_hass(), _make_entry()
        hass.data[DOMAIN] = {}

        await _async_entry_updated(hass, entry)

        hass.config_entries.async_reload.assert_not_called()


class TestCallbackViewRegistration:
    """#20: repeated flow attempts must not stack duplicate routes."""

    def test_views_registered_only_once(self):
        from custom_components.google_outlook_contacts_sync.http import (
            async_register_callback_views,
        )

        hass = _make_hass()
        hass.http.register_view = MagicMock()

        for _ in range(3):
            async_register_callback_views(hass)

        assert hass.http.register_view.call_count == 2  # MS + Google, once each


class TestEnginePersistHook:
    """#7: persistence runs on every path, and never masks the real error."""

    @pytest.mark.asyncio
    async def test_tokens_persisted_after_a_failed_sync(self):
        from custom_components.google_outlook_contacts_sync.sync.engine import (
            SyncEngine,
        )

        calls: list[str] = []

        async def _persist() -> None:
            calls.append("persisted")

        hass = _make_hass()
        google = MagicMock()
        google.fetch_contacts.side_effect = ConnectionError("network down")
        store = MagicMock()
        store.sync_token = "tok"
        store.async_save = AsyncMock()

        engine = SyncEngine(hass, google, MagicMock(), store, persist_tokens=_persist)

        with pytest.raises(ConnectionError):
            await engine.async_sync()

        assert calls == ["persisted"]

    @pytest.mark.asyncio
    async def test_persistence_failure_does_not_mask_sync_error(self):
        from custom_components.google_outlook_contacts_sync.sync.engine import (
            SyncEngine,
        )

        async def _persist() -> None:
            raise RuntimeError("storage exploded")

        hass = _make_hass()
        google = MagicMock()
        google.fetch_contacts.side_effect = ConnectionError("network down")
        store = MagicMock()
        store.sync_token = "tok"
        store.async_save = AsyncMock()

        engine = SyncEngine(hass, google, MagicMock(), store, persist_tokens=_persist)

        # The original network error survives, not the persistence error.
        with pytest.raises(ConnectionError):
            await engine.async_sync()

    @pytest.mark.asyncio
    async def test_dry_run_persists_a_google_refresh(self):
        from custom_components.google_outlook_contacts_sync.sync.engine import (
            SyncEngine,
        )

        calls: list[str] = []

        async def _persist() -> None:
            calls.append("persisted")

        hass = _make_hass()
        google = MagicMock()
        google.fetch_contacts.return_value = ([], "tok")
        store = MagicMock()
        store.sync_token = "tok"
        store.async_save = AsyncMock()

        engine = SyncEngine(hass, google, MagicMock(), store, persist_tokens=_persist)
        await engine.async_plan()

        assert calls == ["persisted"]
