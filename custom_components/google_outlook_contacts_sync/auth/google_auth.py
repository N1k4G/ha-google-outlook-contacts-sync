"""Google OAuth 2.0 helpers for the config flow (redirect-based flow)."""

from __future__ import annotations

import logging
from typing import Any

from google.auth.exceptions import (  # type: ignore[import-untyped]
    RefreshError,
    TransportError,
)
from google.auth.transport.requests import Request  # type: ignore[import-untyped]
from google.oauth2.credentials import Credentials  # type: ignore[import-untyped]
from google_auth_oauthlib.flow import Flow  # type: ignore[import-untyped]

from ..const import GOOGLE_SCOPE
from . import CredentialCheck, CredentialStatus

_LOGGER = logging.getLogger(__name__)

_TOKEN_URI = "https://oauth2.googleapis.com/token"


def build_auth_flow(client_id: str, client_secret: str, redirect_uri: str) -> Flow:
    """Create a Google OAuth flow configured for the HA redirect callback.

    Uses the 'web' client type so Google accepts an arbitrary HTTPS redirect URI.
    The redirect URI must be registered in the Google Cloud Console OAuth client.
    """
    client_config = {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uris": [redirect_uri],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": _TOKEN_URI,
        }
    }
    return Flow.from_client_config(
        client_config,
        scopes=[GOOGLE_SCOPE],
        redirect_uri=redirect_uri,
    )


def get_auth_url(flow: Flow) -> tuple[str, str]:
    """Return (auth_url, state) for the Google OAuth redirect flow."""
    url, state = flow.authorization_url(
        access_type="offline",
        prompt="consent",
        include_granted_scopes="true",
    )
    return str(url), str(state)


def check_credentials(
    client_id: str,
    client_secret: str,
    token_dict: dict[str, Any] | None,
) -> CredentialCheck:
    """Redeem the stored refresh token with the given client credential.

    Returns the refreshed token dict on success. It carries the new client
    credential, which is what the People client is rebuilt from at setup.
    """
    token_dict = token_dict or {}
    if not token_dict.get("refresh_token"):
        return CredentialCheck(CredentialStatus.REAUTHORIZE)

    creds = Credentials(  # type: ignore[no-untyped-call]
        token=None,
        refresh_token=token_dict["refresh_token"],
        token_uri=token_dict.get("token_uri", _TOKEN_URI),
        client_id=client_id,
        client_secret=client_secret,
        scopes=[GOOGLE_SCOPE],
    )
    try:
        creds.refresh(Request())  # type: ignore[no-untyped-call]
    except RefreshError as exc:
        if exc.retryable:
            return CredentialCheck(CredentialStatus.CANNOT_CONNECT)
        error = _oauth_error(exc)
        if error == "invalid_client":
            return CredentialCheck(CredentialStatus.INVALID_CLIENT)
        # invalid_grant (expired or revoked), unauthorized_client (token issued
        # to another client), ...: only a fresh authorization helps.
        _LOGGER.info("Google refresh token rejected (%s)", error or "no error code")
        return CredentialCheck(CredentialStatus.REAUTHORIZE)
    except TransportError as exc:
        _LOGGER.warning("Google credential check failed: %s", type(exc).__name__)
        return CredentialCheck(CredentialStatus.CANNOT_CONNECT)

    return CredentialCheck(
        CredentialStatus.VALID,
        {
            "token": creds.token,
            "refresh_token": creds.refresh_token,
            "token_uri": creds.token_uri,
            "client_id": creds.client_id,
            "client_secret": creds.client_secret,
            "scopes": list(creds.scopes or []),
        },
    )


def _oauth_error(exc: RefreshError) -> str | None:
    """The OAuth error code of a token endpoint rejection, if it carried one."""
    if len(exc.args) > 1 and isinstance(exc.args[1], dict):
        error = exc.args[1].get("error")
        return str(error) if error else None
    return None


def exchange_code(flow: Flow, code: str) -> dict[str, Any]:
    """Exchange the authorization code for credentials and return a token dict."""
    flow.fetch_token(code=code)
    creds = flow.credentials
    return {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes or []),
    }
