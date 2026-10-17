"""Contract tests against the real python-o365 and MSAL code (#38).

The unit tests in test_graph_client.py stub the refresh boundary. These run
GraphContactsClient through the installed O365 Connection and MSAL token cache
with every HTTP request answered by a stub, so an O365 or MSAL upgrade that
changes how a failed refresh surfaces, or how the bearer header is set, fails
here. Nothing reaches the network: an unexpected URL fails the test.
"""

from __future__ import annotations

import json
import time
from collections.abc import Iterator
from typing import Any
from unittest.mock import patch

import pytest
import requests

from custom_components.google_outlook_contacts_sync.auth.ms_auth import (
    build_account,
)
from custom_components.google_outlook_contacts_sync.sync import (
    graph_client as graph_client_module,
)
from custom_components.google_outlook_contacts_sync.sync.graph_client import (
    GraphContactsClient,
    GraphRequestError,
    MSAuthError,
)

_SCOPES = (
    "https://graph.microsoft.com/Contacts.ReadWrite "
    "https://graph.microsoft.com/User.Read"
)
_OPENID = {
    "authorization_endpoint": (
        "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
    ),
    "token_endpoint": "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
    "issuer": (
        "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0"
    ),
}
_INSTANCE = {
    "metadata": [
        {
            "preferred_network": "login.microsoftonline.com",
            "preferred_cache": "login.windows.net",
            "aliases": ["login.microsoftonline.com", "login.windows.net"],
        }
    ],
}
_TOKEN_OK = {
    "token_type": "Bearer",
    "scope": _SCOPES,
    "expires_in": 3600,
    "access_token": "new-at",
    "refresh_token": "new-rt",
    # Entra always returns client_info; MSAL keys cached tokens by it.
    # base64url of {"uid":"uid","utid":"utid"}
    "client_info": "eyJ1aWQiOiJ1aWQiLCJ1dGlkIjoidXRpZCJ9",
}


def _cache(access_token_expires_on: int | str) -> dict[str, Any]:
    """A serialized MSAL cache as the config entry stores it."""
    common = {
        "home_account_id": "uid.utid",
        "environment": "login.microsoftonline.com",
        "client_id": "mid",
    }
    return {
        "Account": {
            "a": {
                "home_account_id": "uid.utid",
                "environment": "login.microsoftonline.com",
                "realm": "consumers",
                "local_account_id": "uid",
                "username": "user@example.com",
                "authority_type": "MSSTS",
            }
        },
        "RefreshToken": {
            "r": {
                **common,
                "credential_type": "RefreshToken",
                "secret": "old-rt",
                "target": f"{_SCOPES} openid profile offline_access",
                "last_modification_time": "1",
            }
        },
        "AccessToken": {
            "t": {
                **common,
                "credential_type": "AccessToken",
                "secret": "old-at",
                "target": _SCOPES,
                "realm": "consumers",
                "token_type": "Bearer",
                "cached_at": str(int(time.time()) - 4000),
                "expires_on": str(access_token_expires_on),
                "extended_expires_on": str(access_token_expires_on),
            }
        },
        "IdToken": {},
        "AppMetadata": {},
    }


def _response(status: int, body: dict[str, Any], url: str) -> requests.Response:
    response = requests.Response()
    response.status_code = status
    response._content = json.dumps(body).encode()
    response.headers["Content-Type"] = "application/json"
    response.url = url
    return response


class StubHttp:
    """Answers Entra and Graph requests; records Graph bearer headers."""

    def __init__(
        self,
        token: tuple[int, dict[str, Any]] = (200, _TOKEN_OK),
        graph: tuple[int, dict[str, Any]] = (201, {"id": "c1"}),
    ) -> None:
        self.token = token
        self.graph = graph
        self.token_requests = 0
        self.graph_auth: list[str | None] = []

    def __call__(
        self, session: requests.Session, method: str, url: str, **kwargs: Any
    ) -> requests.Response:
        if "openid-configuration" in url:
            return _response(200, _OPENID, url)
        if "discovery/instance" in url:
            return _response(200, {**_INSTANCE, "tenant_discovery_endpoint": url}, url)
        if url.endswith("/oauth2/v2.0/token"):
            self.token_requests += 1
            return _response(*self.token, url)
        if url.startswith("https://graph.microsoft.com/"):
            headers = kwargs.get("headers") or {}
            self.graph_auth.append(
                headers.get("Authorization") or session.headers.get("Authorization")
            )
            return _response(*self.graph, url)
        raise AssertionError(f"unexpected request to {url}")


@pytest.fixture
def now() -> int:
    return int(time.time())


def _client(access_token_expires_on: int | str) -> tuple[GraphContactsClient, Any]:
    account, backend = build_account("mid", "secret", _cache(access_token_expires_on))
    assert account.is_authenticated  # what setup does before building the client
    return GraphContactsClient(account), backend


@pytest.fixture
def http() -> Iterator[StubHttp]:
    stub = StubHttp()
    with patch.object(
        requests.Session,
        "request",
        lambda session, method, url, **kw: stub(session, method, url, **kw),
    ):
        yield stub


class TestEnsureTokenAgainstO365:
    def test_fresh_token_makes_no_request(self, http, now) -> None:
        client, _ = _client(now + 3600)

        client.ensure_token()

        assert http.token_requests == 0

    def test_expired_token_is_refreshed_and_kept(self, http, now) -> None:
        client, backend = _client(now - 60)

        client.ensure_token()

        assert http.token_requests == 1
        # The persist hook writes get_token_dict(); it must hold the new tokens.
        assert "new-rt" in json.dumps(backend.get_token_dict())

    @pytest.mark.parametrize("error", ["invalid_client", "invalid_grant"])
    def test_rejection_requires_reauth(self, http, now, error) -> None:
        http.token = (400, {"error": error, "error_description": "AADSTS"})
        client, _ = _client(now - 60)

        with pytest.raises(MSAuthError):
            client.ensure_token()

    def test_unreadable_cache_requires_reauth(self, http) -> None:
        """MSAL's own cache search raises on a corrupted expiry."""
        client, _ = _client("not-a-number")

        with pytest.raises(MSAuthError):
            client.ensure_token()

        assert http.token_requests == 0

    def test_outage_is_not_an_auth_failure(self, http, now) -> None:
        http.token = (503, {"error": "temporarily_unavailable"})
        client, _ = _client(now - 60)

        with pytest.raises(GraphRequestError):
            client.ensure_token()


class TestGraphRequestsAgainstO365:
    def test_new_token_is_sent_without_a_prior_session(self, http, now) -> None:
        client, _ = _client(now - 60)

        client.ensure_token()
        client.create_contact({"givenName": "A"})

        assert http.graph_auth == ["Bearer new-at"]

    def test_new_token_replaces_the_header_of_an_existing_session(
        self, http, now
    ) -> None:
        # Usable now; within the refresh margin by the time the next run starts.
        client, _ = _client(now + 400)
        client.create_contact({"givenName": "A"})  # O365 builds its session

        with patch.object(graph_client_module.time, "time", return_value=now + 200):
            client.ensure_token()
        client.create_contact({"givenName": "B"})

        assert http.graph_auth == ["Bearer old-at", "Bearer new-at"]

    def test_o365_refresh_failure_mid_run_requires_reauth(self, http, now) -> None:
        """Graph answers 401; O365's own refresh fails with a RuntimeError."""
        http.graph = (401, {"error": {"code": "InvalidAuthenticationToken"}})
        http.token = (401, {"error": "invalid_client"})
        # Expired by the clock O365 checks after a 401, but within no margin
        # ensure_token would act on: the run started with a usable token.
        client, _ = _client(now - 1)

        with pytest.raises(MSAuthError):
            client.create_contact({"givenName": "A"})
