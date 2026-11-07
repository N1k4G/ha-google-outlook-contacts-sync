"""Device-code setup and transactional migration at the HA boundary."""

from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from homeassistant.config_entries import SOURCE_REAUTH, SOURCE_RECONFIGURE, SOURCE_USER
from homeassistant.data_entry_flow import FlowResultType
from homeassistant.exceptions import ConfigEntryError

from custom_components.google_outlook_contacts_sync import (
    _async_persist_tokens,
    config_flow,
)
from custom_components.google_outlook_contacts_sync.auth.ms_device import (
    DeviceLoginError,
)
from custom_components.google_outlook_contacts_sync.const import (
    CONF_MS_AUTH_MODE,
    CONF_MS_CLIENT_SECRET,
    DATA_REAUTH_PROVIDER,
    DOMAIN,
    MS_AUTH_CONFIDENTIAL,
    MS_AUTH_DEVICE_CODE,
)
from tests.test_config_flow import (
    GOOGLE_TOKEN,
    MS_TOKEN,
    _make_entry,
    _make_flow,
    _reconfigure_input,
    _written_data,
)
from tests.test_init import _make_entry as _init_entry
from tests.test_init import _make_hass, _setup


@pytest.fixture
def device_login():
    device = MagicMock()
    device.placeholders = {
        "user_code": "TEST-ONLY",
        "verification_uri": "https://microsoft.com/devicelogin",
    }
    device.async_wait = AsyncMock(return_value={"new": "public-cache"})
    with patch.object(config_flow, "start_device_login", return_value=device) as start:
        yield device, start


def _flow(source=SOURCE_RECONFIGURE, entry=None):
    flow = _make_flow(source, entry or _make_entry())
    flow.hass.async_create_background_task.side_effect = (
        lambda coro, name: asyncio.create_task(coro, name=name)
    )
    return flow


async def _finish(flow):
    task = flow._device_task
    assert task is not None
    await asyncio.gather(task, return_exceptions=True)
    done = await flow.async_step_ms_device()
    assert done["type"] is FlowResultType.SHOW_PROGRESS_DONE
    return await flow.async_step_ms_device_complete()


async def _migrate(flow):
    return await flow.async_step_reconfigure(
        _reconfigure_input(**{CONF_MS_AUTH_MODE: MS_AUTH_DEVICE_CODE})
    )


async def test_new_setup_recommends_device_mode():
    flow = _flow(SOURCE_USER)
    result = await flow.async_step_ms_credentials()
    schema = result["data_schema"]
    values = schema({"ms_client_id": "mid"})
    assert values[CONF_MS_AUTH_MODE] == MS_AUTH_DEVICE_CODE
    assert CONF_MS_CLIENT_SECRET not in values


async def test_new_setup_has_no_ms_callback_requirement(device_login):
    flow = _flow(SOURCE_USER)
    flow.hass.config.external_url = None
    result = await flow.async_step_ms_credentials(
        {CONF_MS_AUTH_MODE: MS_AUTH_DEVICE_CODE, "ms_client_id": "mid"}
    )
    assert result["type"] is FlowResultType.SHOW_PROGRESS
    assert set(result["description_placeholders"]) == {"user_code", "verification_uri"}
    result = await _finish(flow)
    assert result["step_id"] == "google_credentials"
    flow._google_client_id = "gid"
    flow._google_client_secret = "synthetic-google-secret"
    flow._google_token = GOOGLE_TOKEN
    created = flow._create_entry()
    assert created["data"][CONF_MS_AUTH_MODE] == MS_AUTH_DEVICE_CODE
    assert CONF_MS_CLIENT_SECRET not in created["data"]


async def test_migration_is_atomic_and_preserves_concurrent_google_refresh(
    device_login,
):
    entry = _make_entry(extra_setting="keep")
    flow = _flow(entry=entry)
    result = await _migrate(flow)
    assert result["type"] is FlowResultType.SHOW_PROGRESS
    assert entry.data[CONF_MS_CLIENT_SECRET] == "msecret"
    flow.hass.config_entries.async_update_entry.assert_not_called()
    refreshed_google = {**GOOGLE_TOKEN, "token": "concurrent-new-google-token"}
    entry.data = {**entry.data, "google_token": refreshed_google}
    result = await _finish(flow)
    assert result["reason"] == "reconfigure_successful"
    written = _written_data(flow)
    assert written[CONF_MS_AUTH_MODE] == MS_AUTH_DEVICE_CODE
    assert written["ms_token"] == {"new": "public-cache"}
    assert CONF_MS_CLIENT_SECRET not in written
    assert written["google_token"] == refreshed_google
    assert written["extra_setting"] == "keep"
    flow.hass.config_entries.async_schedule_reload.assert_called_once_with(
        entry.entry_id
    )


async def test_cancelled_migration_saves_nothing_and_stops_task(device_login):
    device, _ = device_login
    wait = asyncio.Event()
    device.async_wait.side_effect = lambda executor: wait.wait()
    flow = _flow()
    await _migrate(flow)
    task = flow._device_task
    flow.async_remove()
    await asyncio.gather(task, return_exceptions=True)
    device.cancel.assert_called()
    flow.hass.config_entries.async_update_entry.assert_not_called()
    assert (await flow.async_step_ms_device_complete())[
        "reason"
    ] == "ms_device_cancelled"


@pytest.mark.parametrize(
    "reason",
    ["ms_device_denied", "ms_device_expired", "ms_invalid_client", "cannot_connect"],
)
async def test_failed_migration_does_not_modify_existing_entry(device_login, reason):
    device, _ = device_login
    device.async_wait.side_effect = DeviceLoginError(reason)
    flow = _flow()
    await _migrate(flow)
    result = await _finish(flow)
    assert result["step_id"] == "ms_device_failed"
    assert result["errors"] == {"base": reason}
    flow.hass.config_entries.async_update_entry.assert_not_called()
    flow.hass.config_entries.async_schedule_reload.assert_not_called()


async def test_failure_can_explicitly_request_a_new_code(device_login):
    device, start = device_login
    device.async_wait.side_effect = DeviceLoginError("ms_device_expired")
    flow = _flow()
    await _migrate(flow)
    await _finish(flow)
    device.async_wait.side_effect = None
    await flow.async_step_ms_device_failed({})
    await _finish(flow)
    assert start.call_count == 2
    flow._existing_entry.async_on_unload.assert_called_once()
    assert CONF_MS_CLIENT_SECRET not in _written_data(flow)


async def test_public_reauth_uses_device_mode_without_a_secret(device_login):
    entry = _make_entry(**{CONF_MS_AUTH_MODE: MS_AUTH_DEVICE_CODE})
    entry.data.pop(CONF_MS_CLIENT_SECRET)
    flow = _flow(SOURCE_REAUTH, entry)
    flow.hass.data = {DOMAIN: {DATA_REAUTH_PROVIDER: {entry.entry_id: "ms"}}}
    result = await flow.async_step_reauth(dict(entry.data))
    assert result["type"] is FlowResultType.SHOW_PROGRESS
    assert (await _finish(flow))["reason"] == "reauth_successful"
    assert CONF_MS_CLIENT_SECRET not in _written_data(flow)


async def test_late_initiation_cannot_start_polling_after_removal(device_login):
    device, _ = device_login
    started, finish = asyncio.Event(), asyncio.Event()
    flow = _flow()
    flow._ms_client_id = "mid"

    async def delayed(*args):
        started.set()
        await finish.wait()
        return device

    flow.hass.async_add_executor_job = delayed
    task = asyncio.create_task(flow.async_step_ms_device())
    await started.wait()
    flow.async_remove()
    finish.set()
    assert (await task)["reason"] == "ms_device_cancelled"
    flow.hass.async_create_background_task.assert_not_called()
    flow.hass.config_entries.async_update_entry.assert_not_called()
    device.cancel.assert_called_once()


@pytest.mark.parametrize("stop", ["shutdown", "unload"])
async def test_ha_shutdown_or_entry_unload_stops_pending_login(device_login, stop):
    device, _ = device_login
    wait = asyncio.Event()
    device.async_wait.side_effect = lambda executor: wait.wait()
    flow = _flow()
    await _migrate(flow)
    task = flow._device_task
    if stop == "shutdown":
        callback = flow.hass.bus.async_listen_once.call_args.args[1]
        callback(MagicMock())
    else:
        callback = flow._existing_entry.async_on_unload.call_args.args[0]
        callback()
    await asyncio.gather(task, return_exceptions=True)
    assert flow._removed
    flow.hass.config_entries.async_update_entry.assert_not_called()


async def test_unknown_mode_in_setup_and_reconfigure_is_rejected():
    flow = _flow()
    result = await flow.async_step_reconfigure(
        _reconfigure_input(**{CONF_MS_AUTH_MODE: "invalid"})
    )
    assert result["errors"] == {"base": "ms_invalid_mode"}
    flow.hass.config_entries.async_update_entry.assert_not_called()
    entry = _make_entry(**{CONF_MS_AUTH_MODE: "invalid"})
    with pytest.raises(ConfigEntryError, match="Unknown"):
        await _setup(_make_hass(), entry)


async def test_public_entry_loads_without_a_secret_after_restart():
    entry = _init_entry()
    entry.data[CONF_MS_AUTH_MODE] = MS_AUTH_DEVICE_CODE
    entry.data.pop(CONF_MS_CLIENT_SECRET)
    result, _ = await _setup(_make_hass(), entry)
    assert result is True


async def test_legacy_reconfigure_keeps_mode_and_requires_new_confidential_secret():
    flow = _flow()
    result = await flow.async_step_reconfigure(_reconfigure_input())
    assert result["reason"] == "reconfigure_successful"
    assert CONF_MS_AUTH_MODE not in _written_data(flow)
    entry = _make_entry(**{CONF_MS_AUTH_MODE: MS_AUTH_DEVICE_CODE})
    entry.data.pop(CONF_MS_CLIENT_SECRET)
    flow = _flow(entry=entry)
    result = await flow.async_step_reconfigure(
        _reconfigure_input(**{CONF_MS_AUTH_MODE: MS_AUTH_CONFIDENTIAL})
    )
    assert result["errors"] == {"base": "new_client_requires_secret"}


async def test_old_runtime_cannot_restore_tokens_after_migration():
    entry = _make_entry(**{CONF_MS_AUTH_MODE: MS_AUTH_DEVICE_CODE})
    entry.data.pop(CONF_MS_CLIENT_SECRET)
    entry.data["ms_token"] = {"new": "public-cache"}
    hass = _make_hass()
    backend = MagicMock()
    backend.get_token_dict.return_value = {"old": "confidential-cache"}
    google = MagicMock()
    google.get_token_dict.return_value = GOOGLE_TOKEN
    await _async_persist_tokens(
        hass,
        entry,
        backend,
        google,
        {"ms_token": MS_TOKEN, "google_token": GOOGLE_TOKEN},
    )
    hass.config_entries.async_update_entry.assert_not_called()


@pytest.mark.parametrize("same_account", [True, False])
async def test_migration_must_preserve_microsoft_account_identity(
    device_login, same_account
):
    device, _ = device_login
    entry = _make_entry(
        ms_token={
            **MS_TOKEN,
            "Account": {"old": {"home_account_id": "original-account"}},
        }
    )
    device.async_wait.return_value = {
        "Account": {
            "new": {
                "home_account_id": "original-account"
                if same_account
                else "other-account"
            }
        },
        "RefreshToken": {"r": {"secret": "synthetic-new-rt"}},
    }
    flow = _flow(entry=entry)
    await _migrate(flow)
    result = await _finish(flow)
    if same_account:
        assert result["reason"] == "reconfigure_successful"
    else:
        assert result["errors"] == {"base": "ms_device_account_mismatch"}
        flow.hass.config_entries.async_update_entry.assert_not_called()


async def test_shutdown_does_not_remove_the_one_shot_listener_twice(device_login):
    device, _ = device_login
    device.async_wait.side_effect = lambda executor: asyncio.Event().wait()
    flow = _flow()
    await _migrate(flow)
    task = flow._device_task
    remove_listener = flow.hass.bus.async_listen_once.return_value
    stop = flow.hass.bus.async_listen_once.call_args.args[1]
    stop(MagicMock())
    await asyncio.gather(task, return_exceptions=True)
    remove_listener.assert_not_called()
