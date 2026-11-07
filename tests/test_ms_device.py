"""Device-code lifecycle and real-SDK contracts with no external HTTP."""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import time
from unittest.mock import MagicMock, patch

import pytest
import requests

from custom_components.google_outlook_contacts_sync.auth import (
    CredentialStatus,
    ms_device,
)
from custom_components.google_outlook_contacts_sync.auth.ms_auth import (
    HAStoreTokenBackend,
    build_account,
)
from custom_components.google_outlook_contacts_sync.auth.ms_device import (
    DeviceAuthorization,
    DeviceLoginError,
    start_device_login,
)
from custom_components.google_outlook_contacts_sync.const import MS_AUTH_DEVICE_CODE
from custom_components.google_outlook_contacts_sync.sync.graph_client import (
    GraphContactsClient,
)
from tests.test_o365_contract import _OPENID, _TOKEN_OK, StubHttp, _response

_DEVICE_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode"


def _flow(**changes):
    return {
        "user_code": "TEST-ONLY",
        "device_code": "synthetic-device-code",
        "verification_uri": "https://microsoft.com/devicelogin",
        "expires_in": 900,
        "interval": 5,
        **changes,
    }


async def _execute(fn, *args):
    return fn(*args)


@pytest.fixture
def device(monkeypatch):
    backend = MagicMock(spec=HAStoreTokenBackend)
    backend.get_token_dict.return_value = {"verified": "synthetic-cache"}
    account = MagicMock()
    account.con.msal_client.acquire_token_by_device_flow.return_value = {
        "access_token": "synthetic-access"
    }
    monkeypatch.setattr(
        ms_device,
        "refresh_access_token",
        lambda *_: (CredentialStatus.VALID, "synthetic-refreshed"),
    )
    return DeviceAuthorization(account, backend, _flow())


@pytest.fixture
def sleeps(monkeypatch):
    delays = []
    real_sleep = asyncio.sleep

    async def immediate(delay):
        delays.append(delay)
        await real_sleep(0)

    monkeypatch.setattr(ms_device, "async_sleep", immediate)
    return delays


@pytest.mark.parametrize(
    "changes",
    [
        {"verification_uri": "https://evil.example/devicelogin"},
        {"verification_uri": "http://microsoft.com/devicelogin"},
        {"verification_uri": ["unexpected"]},
        {"user_code": "[click](https://evil.example)"},
        {"user_code": ""},
        {"expires_in": 0},
        {"expires_in": None},
        {"interval": float("nan")},
        {"interval": -1},
        {"device_code": ""},
    ],
)
def test_invalid_provider_flow_is_not_displayed(changes):
    with pytest.raises(DeviceLoginError, match="ms_device_failed"):
        DeviceAuthorization(MagicMock(), MagicMock(), _flow(**changes))


def test_no_device_code_or_cache_in_repr_or_placeholders(device):
    assert set(device.placeholders) == {"user_code", "verification_uri"}
    assert "synthetic-device-code" not in repr(device)
    assert "synthetic-cache" not in repr(device)


async def test_success_requires_verified_refresh(device, sleeps):
    assert await device.async_wait(_execute) == {"verified": "synthetic-cache"}
    assert sleeps == [5]
    call = device.account.con.msal_client.acquire_token_by_device_flow.call_args
    assert call.kwargs["exit_condition"]({}) is True
    device.backend.get_token_dict.assert_called_once()


@pytest.mark.parametrize(
    ("error", "reason"),
    [
        ("authorization_declined", "ms_device_denied"),
        ("access_denied", "ms_device_denied"),
        ("expired_token", "ms_device_expired"),
        ("bad_verification_code", "ms_device_failed"),
        ("invalid_client", "ms_invalid_client"),
        ("unauthorized_client", "ms_invalid_client"),
    ],
)
async def test_terminal_provider_errors_save_nothing(device, sleeps, error, reason):
    device.account.con.msal_client.acquire_token_by_device_flow.return_value = {
        "error": error,
        "error_description": "synthetic-sensitive-detail",
    }
    with pytest.raises(DeviceLoginError, match=reason) as exc:
        await device.async_wait(_execute)
    assert "synthetic-sensitive-detail" not in str(exc.value)
    device.backend.get_token_dict.assert_not_called()


async def test_pending_and_slow_down_respect_returned_interval(device, sleeps):
    attempts = 0

    def poll(*_args, **_kwargs):
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            return {"error": "authorization_pending"}
        if attempts == 2:
            device.flow["interval"] = 10
            return {"error": "slow_down"}
        return {"access_token": "synthetic-access"}

    device.account.con.msal_client.acquire_token_by_device_flow.side_effect = poll
    await device.async_wait(_execute)
    assert sleeps == [5, 5, 10]


@pytest.mark.parametrize("provider_error", [False, True])
async def test_network_and_server_failures_are_bounded(device, sleeps, provider_error):
    poll = device.account.con.msal_client.acquire_token_by_device_flow
    if provider_error:
        poll.return_value = {"error": "temporarily_unavailable"}
    else:
        poll.side_effect = ConnectionError("synthetic-sensitive-detail")
    with pytest.raises(DeviceLoginError, match="cannot_connect"):
        await device.async_wait(_execute)
    assert poll.call_count == 3
    assert sleeps == [5, 10, 20]


async def test_expiry_stops_without_token_exchange(device, sleeps):
    device._deadline = time.monotonic() - 1
    with pytest.raises(DeviceLoginError, match="ms_device_expired"):
        await device.async_wait(_execute)
    device.account.con.msal_client.acquire_token_by_device_flow.assert_not_called()


async def test_cancellation_discards_late_http_result(device, sleeps):
    started, finish = asyncio.Event(), asyncio.Event()

    async def delayed(fn):
        started.set()
        await finish.wait()
        return {"access_token": "synthetic-late-access"}

    task = asyncio.create_task(device.async_wait(delayed))
    await started.wait()
    device.cancel()
    finish.set()
    with pytest.raises(DeviceLoginError, match="ms_device_cancelled"):
        await task
    device.backend.get_token_dict.assert_not_called()


async def test_task_cancellation_stops_future_polls(device, sleeps):
    started = asyncio.Event()

    async def pending(fn):
        started.set()
        await asyncio.Event().wait()

    task = asyncio.create_task(device.async_wait(pending))
    await started.wait()
    task.cancel()
    with pytest.raises(asyncio.CancelledError):
        await task
    assert device._cancelled.is_set()
    device.backend.get_token_dict.assert_not_called()


@pytest.mark.parametrize(
    ("status", "reason"),
    [
        (CredentialStatus.INVALID_CLIENT, "ms_invalid_client"),
        (CredentialStatus.REAUTHORIZE, "ms_device_failed"),
        (CredentialStatus.CANNOT_CONNECT, "cannot_connect"),
    ],
)
async def test_unrefreshable_new_login_is_not_accepted(
    device, sleeps, monkeypatch, status, reason
):
    monkeypatch.setattr(ms_device, "refresh_access_token", lambda *_: (status, None))
    with pytest.raises(DeviceLoginError, match=reason):
        await device.async_wait(_execute)
    device.backend.get_token_dict.assert_not_called()


def test_public_factory_ignores_any_legacy_secret():
    account, _ = build_account("synthetic-id", "ignored", None, MS_AUTH_DEVICE_CODE)
    assert account.con.auth == ("synthetic-id",)
    assert account.con.auth_flow_type == "public"
    assert account.con.timeout == (5, 10)


def test_unknown_mode_is_rejected():
    with pytest.raises(ValueError, match="Unknown"):
        build_account("synthetic-id", "", None, "unsupported")


class DeviceHttp(StubHttp):
    def __init__(self):
        super().__init__()
        self.posts = []
        claims = {
            "aud": "mid",
            "iss": _OPENID["issuer"],
            "sub": "synthetic-subject",
            "iat": int(time.time()),
            "exp": int(time.time()) + 3600,
            "tid": "utid",
            "oid": "uid",
            "preferred_username": "user@example.com",
        }

        def encode(data):
            return (
                base64.urlsafe_b64encode(json.dumps(data).encode())
                .rstrip(b"=")
                .decode()
            )

        # Synthetic ID token; the HTTP boundary is trusted only inside this stub.
        id_token = encode({"alg": "none"}) + "." + encode(claims) + "."
        self.token = (200, {**_TOKEN_OK, "id_token": id_token})

    def __call__(self, session, method, url, **kwargs):
        if "openid-configuration" in url:
            return _response(
                200, {**_OPENID, "device_authorization_endpoint": _DEVICE_URL}, url
            )
        if url == _DEVICE_URL:
            return _response(200, _flow(), url)
        if url.endswith("/token"):
            self.posts.append(dict(kwargs.get("data") or {}))
        return super().__call__(session, method, url, **kwargs)


async def test_real_sdks_device_login_refresh_restart_and_graph(sleeps, caplog):
    """Actual MSAL/O365, synthetic HTTP: no account or contact is accessed."""
    http = DeviceHttp()
    caplog.set_level(logging.DEBUG)
    with patch.object(
        requests.Session,
        "request",
        lambda session, method, url, **kw: http(session, method, url, **kw),
    ):
        device = start_device_login("mid")
        cache = await device.async_wait(_execute)
        assert "new-rt" in json.dumps(cache)
        restarted, _ = build_account("mid", "", cache, MS_AUTH_DEVICE_CODE)
        assert restarted.is_authenticated
        client = GraphContactsClient(restarted)
        client.ensure_token()
        client.create_contact({"givenName": "Synthetic"})
    assert http.graph_auth == ["Bearer new-at"]
    assert [post["grant_type"] for post in http.posts] == [
        "urn:ietf:params:oauth:grant-type:device_code",
        "refresh_token",
    ]
    for post in http.posts:
        assert "client_secret" not in post
        assert "client_assertion" not in post
        assert "redirect_uri" not in post

    assert "synthetic-device-code" not in caplog.text
    assert "new-rt" not in caplog.text
    assert "new-at" not in caplog.text
    assert "TEST-ONLY" not in caplog.text
