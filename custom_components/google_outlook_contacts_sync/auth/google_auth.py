"""Google OAuth 2.0 helpers for the config flow (redirect-based flow)."""

from __future__ import annotations

from typing import Any

from google_auth_oauthlib.flow import Flow  # type: ignore[import-untyped]

from ..const import GOOGLE_SCOPE


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
            "token_uri": "https://oauth2.googleapis.com/token",
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
