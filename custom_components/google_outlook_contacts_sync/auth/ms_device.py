"""Bounded Microsoft device authorization; never writes a HA config entry."""

from __future__ import annotations

import asyncio
import math
import re
import time
from asyncio import sleep as async_sleep
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from threading import Event
from typing import Any

from ..const import MS_AUTH_DEVICE_CODE, MS_SCOPES
from . import CredentialStatus
from .ms_auth import HAStoreTokenBackend, build_account, refresh_access_token

# Only trusted Microsoft destinations may be displayed as a sign-in link.
_VERIFICATION_URIS = {
    "https://microsoft.com/devicelogin",
    "https://www.microsoft.com/link",
}
_PENDING = {"authorization_pending", "slow_down"}


class DeviceLoginError(Exception):
    """A translated, non-sensitive failure reason suitable for the config flow."""


def _reason(error: object) -> str:
    if error in ("authorization_declined", "access_denied"):
        return "ms_device_denied"
    if error in ("expired_token", "code_expired"):
        return "ms_device_expired"
    if error in ("invalid_client", "unauthorized_client"):
        return "ms_invalid_client"
    return "ms_device_failed"


@dataclass(repr=False)
class DeviceAuthorization:
    """An ephemeral SDK flow; cancelled workers can only mutate this object."""

    account: Any
    backend: HAStoreTokenBackend
    flow: dict[str, Any]
    _cancelled: Event = field(default_factory=Event)
    _deadline: float = field(init=False)

    def __post_init__(self) -> None:
        try:
            expires = float(self.flow.get("expires_in", 0))
            interval = float(self.flow.get("interval", 5))
        except (TypeError, ValueError):
            raise DeviceLoginError("ms_device_failed") from None
        if (
            not math.isfinite(expires)
            or not math.isfinite(interval)
            or expires <= 0
            or interval <= 0
            or not isinstance(self.flow.get("verification_uri"), str)
            or self.flow["verification_uri"] not in _VERIFICATION_URIS
            or not isinstance(self.flow.get("user_code"), str)
            or re.fullmatch(r"[A-Za-z0-9-]{1,32}", self.flow["user_code"]) is None
            or not self.flow.get("device_code")
        ):
            raise DeviceLoginError("ms_device_failed")
        self._deadline = time.monotonic() + min(expires, 900)

    @property
    def placeholders(self) -> dict[str, str]:
        """Only the user-facing code and trusted link, never the device code."""
        return {
            "user_code": self.flow["user_code"],
            "verification_uri": self.flow["verification_uri"],
        }

    def cancel(self) -> None:
        """Stop scheduling work, including when an executor request is in flight."""
        self._cancelled.set()

    def _poll_once(self) -> dict[str, Any]:
        if self._cancelled.is_set():
            return {"error": "cancelled"}
        result: dict[str, Any] = (
            self.account.con.msal_client.acquire_token_by_device_flow(
                self.flow, exit_condition=lambda _flow: True
            )
        )
        return result

    def _verify(self) -> dict[str, Any]:
        if self._cancelled.is_set():
            raise DeviceLoginError("ms_device_cancelled")
        status, _ = refresh_access_token(self.account, self.backend)
        if status is CredentialStatus.CANNOT_CONNECT:
            raise DeviceLoginError("cannot_connect")
        if status is CredentialStatus.INVALID_CLIENT:
            raise DeviceLoginError("ms_invalid_client")
        if status is not CredentialStatus.VALID:
            raise DeviceLoginError("ms_device_failed")
        return self.backend.get_token_dict()

    async def async_wait(
        self, executor: Callable[..., Awaitable[Any]]
    ) -> dict[str, Any]:
        """Poll once per interval, with a deadline and bounded network retries."""
        failures = 0
        try:
            async with asyncio.timeout(max(0, self._deadline - time.monotonic())):
                while not self._cancelled.is_set():
                    interval = float(self.flow.get("interval", 5))
                    if not math.isfinite(interval) or interval <= 0:
                        raise DeviceLoginError("ms_device_failed")
                    delay = max(interval, min(interval * (2**failures), 60))
                    await async_sleep(delay)
                    if self._cancelled.is_set():
                        break
                    try:
                        result = await executor(self._poll_once)
                    except Exception:
                        failures += 1
                        if failures >= 3:
                            raise DeviceLoginError("cannot_connect") from None
                        continue
                    if self._cancelled.is_set():
                        break
                    error = result.get("error")
                    if error in ("temporarily_unavailable", "server_error"):
                        failures += 1
                        if failures >= 3:
                            raise DeviceLoginError("cannot_connect")
                        continue
                    failures = 0
                    if "access_token" in result:
                        try:
                            token = await executor(self._verify)
                        except DeviceLoginError:
                            raise
                        except Exception:
                            raise DeviceLoginError("cannot_connect") from None
                        if self._cancelled.is_set():
                            break
                        return dict(token)
                    if error in _PENDING:
                        continue
                    raise DeviceLoginError(_reason(error))
        except TimeoutError:
            raise DeviceLoginError("ms_device_expired") from None
        finally:
            self.cancel()
        raise DeviceLoginError("ms_device_cancelled")


def start_device_login(client_id: str) -> DeviceAuthorization:
    """Initiate through MSAL in an executor; omit all client credentials."""
    account, backend = build_account(client_id, "", None, MS_AUTH_DEVICE_CODE)
    flow = account.con.msal_client.initiate_device_flow(scopes=MS_SCOPES)
    if "error" in flow:
        raise DeviceLoginError(_reason(flow.get("error")))
    return DeviceAuthorization(account, backend, flow)
