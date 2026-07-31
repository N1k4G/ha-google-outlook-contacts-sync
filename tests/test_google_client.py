"""Tests for sync/google_client.py — the adapter boundary to the People API."""

from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock, patch

import pytest
from google.auth.exceptions import RefreshError
from googleapiclient.errors import HttpError

from custom_components.google_outlook_contacts_sync.sync.google_client import (
    GoogleAuthError,
    GoogleContactsClient,
    SyncTokenExpiredError,
)


class FakeRequest:
    def __init__(self, page: Any) -> None:
        self._page = page

    def execute(self) -> Any:
        if isinstance(self._page, Exception):
            raise self._page
        return self._page


class FakeConnections:
    """Records list() kwargs and replays queued pages."""

    def __init__(self, pages: list[Any]) -> None:
        self._pages = list(pages)
        self.calls: list[dict[str, Any]] = []

    def list(self, **kwargs: Any) -> FakeRequest:
        self.calls.append(kwargs)
        if not self._pages:
            raise AssertionError("unexpected extra People API page request")
        return FakeRequest(self._pages.pop(0))


def _client(pages: list[Any]) -> tuple[GoogleContactsClient, FakeConnections]:
    connections = FakeConnections(pages)
    service = MagicMock()
    service.people.return_value.connections.return_value = connections
    client = GoogleContactsClient(MagicMock())
    with patch(
        "custom_components.google_outlook_contacts_sync.sync.google_client.build",
        return_value=service,
    ):
        client._get_service()
    return client, connections


def _http_error(status: int, content: bytes = b"") -> HttpError:
    resp = MagicMock()
    resp.status = status
    resp.reason = "error"
    return HttpError(resp, content)


class TestFetchContacts:
    def test_first_full_fetch_requests_a_sync_token(self):
        client, connections = _client([{"connections": [], "nextSyncToken": "tok"}])

        people, token = client.fetch_contacts(None)

        assert people == []
        assert token == "tok"
        assert connections.calls[0]["requestSyncToken"] is True
        assert "syncToken" not in connections.calls[0]

    def test_delta_fetch_sends_the_stored_cursor(self):
        client, connections = _client([{"connections": [], "nextSyncToken": "new"}])

        client.fetch_contacts("old-token")

        assert connections.calls[0]["syncToken"] == "old-token"

    def test_paging_combines_pages_and_forwards_the_page_token(self):
        client, connections = _client(
            [
                {"connections": [{"resourceName": "people/a"}], "nextPageToken": "p2"},
                {
                    "connections": [{"resourceName": "people/b"}],
                    "nextSyncToken": "final",
                },
            ]
        )

        people, token = client.fetch_contacts(None)

        assert [p["resourceName"] for p in people] == ["people/a", "people/b"]
        # The sync token comes from the last page, not the first.
        assert token == "final"
        assert "pageToken" not in connections.calls[0]
        assert connections.calls[1]["pageToken"] == "p2"

    def test_missing_connections_key_is_treated_as_empty(self):
        client, _ = _client([{"nextSyncToken": "tok"}])
        people, _ = client.fetch_contacts(None)
        assert people == []


class TestErrorClassification:
    def test_410_means_the_sync_token_expired(self):
        client, _ = _client([_http_error(410)])
        with pytest.raises(SyncTokenExpiredError):
            client.fetch_contacts("stale")

    def test_400_with_expired_marker_means_the_sync_token_expired(self):
        client, _ = _client([_http_error(400, b'{"error":"EXPIRED_SYNC_TOKEN"}')])
        with pytest.raises(SyncTokenExpiredError):
            client.fetch_contacts("stale")

    def test_plain_400_is_not_mistaken_for_an_expired_token(self):
        """Misreading this would discard a perfectly good cursor."""
        client, _ = _client([_http_error(400, b'{"error":"badRequest"}')])
        with pytest.raises(HttpError):
            client.fetch_contacts("fine")

    def test_500_is_not_mistaken_for_an_expired_token(self):
        client, _ = _client([_http_error(500, b"server error")])
        with pytest.raises(HttpError):
            client.fetch_contacts("fine")

    def test_refresh_error_becomes_google_auth_error(self):
        """GoogleAuthError is what routes the entry into the reauth flow."""
        client, _ = _client([RefreshError("invalid_grant")])
        with pytest.raises(GoogleAuthError):
            client.fetch_contacts(None)


class TestTokenSerialization:
    def test_roundtrip_preserves_the_credential_fields(self):
        token = {
            "token": "access",
            "refresh_token": "refresh",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_id": "cid",
            "client_secret": "csecret",
        }
        client = GoogleContactsClient.from_token_dict(token)
        result = client.get_token_dict()

        assert result["refresh_token"] == "refresh"
        assert result["client_id"] == "cid"
        assert result["client_secret"] == "csecret"

    def test_missing_keys_do_not_raise(self):
        """#16: an incomplete token must surface as reauth, not KeyError."""
        client = GoogleContactsClient.from_token_dict({})
        assert client.get_token_dict()["client_id"] is None

    def test_serializes_a_token_refreshed_at_runtime(self):
        """#7: the refreshed access token must be what gets persisted."""
        token = {
            "token": "old",
            "refresh_token": "refresh",
            "client_id": "cid",
            "client_secret": "csecret",
        }
        client = GoogleContactsClient.from_token_dict(token)
        client._credentials.token = "refreshed-by-the-sdk"

        assert client.get_token_dict()["token"] == "refreshed-by-the-sdk"
