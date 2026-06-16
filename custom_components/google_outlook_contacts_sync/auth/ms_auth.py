"""Microsoft OAuth helpers using python-o365.

The "alternate auth" flow:
1. Config flow calls ``get_auth_url`` to get the MS sign-in URL.
2. The URL embeds the HA config flow_id in the OAuth *state* parameter so
   the callback view can identify which flow to resume.
3. The HA callback view (http.py) receives the redirect, extracts the code,
   and calls ``request_token`` with the full redirect URL.
4. python-o365 exchanges the code for tokens and persists them via the
   token backend.
"""

from __future__ import annotations

import json
import logging
from typing import Any

from O365 import Account  # type: ignore[import-untyped]
from O365.utils.token import BaseTokenBackend  # type: ignore[import-untyped]

from ..const import MS_AUTH_CALLBACK_PATH, MS_SCOPES

_LOGGER = logging.getLogger(__name__)


class HAStoreTokenBackend(BaseTokenBackend):  # type: ignore[misc]
    """Token backend that persists to an in-memory dict (caller saves to HA Store).

    BaseTokenBackend inherits from MSAL's TokenCache. The live token data lives
    in the inherited self._cache dict; self.serialize()/self.deserialize() convert
    it to/from JSON. self._token is our own HA-facing copy kept in sync on every
    save/load.
    """

    def __init__(self, token_dict: dict[str, Any] | None = None) -> None:
        super().__init__()
        self._token: dict[str, Any] = token_dict or {}

    def load_token(self) -> bool:
        if not self._token:
            return False
        try:
            # O365's BaseTokenBackend.deserialize() only returns the parsed dict;
            # it does NOT update self._cache. We must update self._cache directly
            # so that MSAL can find the tokens for is_authenticated checks.
            with self._lock:
                self._cache.update(self._token)
            return True
        except Exception:
            _LOGGER.exception("HAStoreTokenBackend: failed to load token")
            return False

    def save_token(self, force: bool = False) -> bool:  # noqa: FBT001
        try:
            self._token = json.loads(self.serialize())
            return True
        except Exception:
            _LOGGER.exception("HAStoreTokenBackend: failed to serialize token")
            return False

    def delete_token(self) -> bool:
        self._token = {}
        with self._lock:
            self._cache.clear()
        return True

    def check_token(self) -> bool:
        return bool(self._token)

    def get_token_dict(self) -> dict[str, Any]:
        """Return the current token as a plain dict for HA storage."""
        return dict(self._token)


def build_account(
    client_id: str,
    client_secret: str,
    token_dict: dict[str, Any] | None = None,
) -> tuple[Account, HAStoreTokenBackend]:
    """Create an O365 Account and the associated token backend."""
    backend = HAStoreTokenBackend(token_dict)
    account = Account(
        (client_id, client_secret),
        token_backend=backend,
        tenant_id="consumers",
    )
    return account, backend


def get_auth_url(
    account: Account,
    redirect_uri: str,
) -> tuple[str, dict[str, Any]]:
    """Return (auth_url, msal_flow) for the Microsoft OAuth redirect flow.

    O365 returns the full MSAL auth-code flow dict as the second element.
    The caller extracts flow["state"] for the state→flow_id mapping in
    hass.data and passes the whole dict to request_token later.
    """
    url, flow = account.con.get_authorization_url(
        requested_scopes=MS_SCOPES,
        redirect_uri=redirect_uri,
    )
    return url, flow if isinstance(flow, dict) else {"state": str(flow)}


def request_token(
    account: Account,
    authorization_response_url: str,
    ms_flow: dict[str, Any],
) -> bool:
    """Exchange the callback URL for tokens. Returns True on success."""
    try:
        result: bool = account.con.request_token(
            authorization_response_url,
            flow=ms_flow,
        )
        return result
    except Exception:
        _LOGGER.exception("Failed to exchange MS OAuth code for tokens")
        return False


def build_redirect_uri(ha_external_url: str) -> str:
    """Build the full redirect URI from the HA external URL."""
    return ha_external_url.rstrip("/") + MS_AUTH_CALLBACK_PATH
