"""Tests for sync/graph_client.py — the adapter boundary to O365/requests.

These tests deliberately avoid ``MagicMock`` for API responses. ``O365.Connection``
returns a ``requests.Response``, which offers ``.json()`` but *not* ``.get()``.
A MagicMock would answer both and silently hide exactly the class of bug this
module had (calling ``.get("id")`` on the response object).
"""

from __future__ import annotations

import json
import time
from unittest.mock import MagicMock, patch

import pytest
import requests

from custom_components.google_outlook_contacts_sync.auth import CredentialStatus
from custom_components.google_outlook_contacts_sync.sync import graph_client
from custom_components.google_outlook_contacts_sync.sync.graph_client import (
    GraphContactsClient,
    GraphRequestError,
    MSAuthError,
)


class FakeResponse:
    """Minimal stand-in for requests.Response.

    Exposes only the surface the client is allowed to use: truthiness and
    ``.json()``. Accessing anything else raises AttributeError, which is the
    point — it keeps the adapter honest.
    """

    def __init__(self, body=None, ok=True, json_error=None):
        self._body = body
        self._ok = ok
        self._json_error = json_error

    def __bool__(self):
        return self._ok

    def json(self):
        if self._json_error is not None:
            raise self._json_error
        return self._body


class FakeConnection:
    """Records requests and replays queued responses."""

    def __init__(self, get=None, post=None, patch=None, delete=None):
        self._get_queue = list(get or [])
        self._post = post
        self._patch = patch
        self._delete = delete
        self.get_calls: list[tuple[str, dict]] = []
        self.post_calls: list[tuple[str, dict]] = []
        self.patch_calls: list[tuple[str, dict]] = []
        self.delete_calls: list[str] = []

    def get(self, url, params=None):
        self.get_calls.append((url, params))
        if not self._get_queue:
            raise AssertionError(f"unexpected extra GET for {url}")
        return self._get_queue.pop(0)

    def post(self, url, data=None):
        self.post_calls.append((url, data))
        return self._post

    def patch(self, url, data=None):
        self.patch_calls.append((url, data))
        return self._patch

    def delete(self, url):
        self.delete_calls.append(url)
        return self._delete


def _http_error(status: int, body: object) -> requests.HTTPError:
    """The HTTPError O365 raises for a 4xx/5xx, carrying a real Response."""
    response = requests.Response()
    response.status_code = status
    response._content = body if isinstance(body, bytes) else json.dumps(body).encode()
    return requests.HTTPError(
        f"{status} Server Error | Error Message: x", response=response
    )


class RaisingConnection(FakeConnection):
    """A connection whose PATCH raises, as O365 does for an error status."""

    def __init__(self, error: Exception) -> None:
        super().__init__()
        self._error = error

    def patch(self, url, data=None):
        self.patch_calls.append((url, data))
        raise self._error


class FakeAccount:
    def __init__(self, con, is_authenticated=True):
        self.con = con
        self.is_authenticated = is_authenticated


def _client(con, *, authenticated=True) -> GraphContactsClient:
    return GraphContactsClient(FakeAccount(con, is_authenticated=authenticated))


class TestCreateContact:
    def test_returns_id_from_json_body(self):
        con = FakeConnection(post=FakeResponse({"id": "graph-001"}))
        assert _client(con).create_contact({"givenName": "A"}) == "graph-001"
        assert con.post_calls[0][1] == {"givenName": "A"}

    def test_response_object_is_not_treated_as_a_dict(self):
        """Regression for #6: the old code called .get() on the Response."""
        response = FakeResponse({"id": "graph-001"})
        with pytest.raises(AttributeError):
            response.get("id")  # type: ignore[attr-defined]
        assert _client(FakeConnection(post=response)).create_contact({}) == "graph-001"

    def test_missing_id_raises(self):
        con = FakeConnection(post=FakeResponse({"displayName": "no id here"}))
        with pytest.raises(GraphRequestError, match="no contact id"):
            _client(con).create_contact({})

    def test_empty_id_raises(self):
        con = FakeConnection(post=FakeResponse({"id": ""}))
        with pytest.raises(GraphRequestError, match="no contact id"):
            _client(con).create_contact({})

    def test_failed_post_while_authenticated_raises_request_error(self):
        con = FakeConnection(post=FakeResponse(ok=False))
        with pytest.raises(GraphRequestError, match="creation"):
            _client(con).create_contact({})

    def test_failed_post_while_unauthenticated_raises_auth_error(self):
        con = FakeConnection(post=FakeResponse(ok=False))
        with pytest.raises(MSAuthError):
            _client(con, authenticated=False).create_contact({})

    def test_unparsable_body_raises(self):
        con = FakeConnection(post=FakeResponse(json_error=ValueError("not json")))
        with pytest.raises(GraphRequestError, match="unparsable"):
            _client(con).create_contact({})


class TestUpdateContact:
    def test_success_sends_body_to_contact_url(self):
        con = FakeConnection(patch=FakeResponse({}))
        _client(con).update_contact("graph-1", {"surname": "B"})
        url, data = con.patch_calls[0]
        assert url.endswith("/me/contacts/graph-1")
        assert data == {"surname": "B"}

    def test_failed_patch_while_authenticated_raises_request_error(self):
        con = FakeConnection(patch=FakeResponse(ok=False))
        with pytest.raises(GraphRequestError, match="update"):
            _client(con).update_contact("graph-1", {})

    def test_failed_patch_while_unauthenticated_raises_auth_error(self):
        con = FakeConnection(patch=FakeResponse(ok=False))
        with pytest.raises(MSAuthError):
            _client(con, authenticated=False).update_contact("graph-1", {})

    def test_http_error_reports_status_and_graph_error_code(self):
        error = _http_error(
            500,
            {
                "error": {
                    "code": "ErrorInternalServerError",
                    "message": "An internal server error occurred.",
                }
            },
        )
        with pytest.raises(GraphRequestError) as excinfo:
            _client(RaisingConnection(error)).update_contact("graph-1", {})
        assert str(excinfo.value) == (
            "Microsoft Graph PATCH failed: HTTP 500 ErrorInternalServerError: "
            "An internal server error occurred."
        )
        assert excinfo.value.__cause__ is error

    def test_http_error_with_unparsable_body_still_reports_status(self):
        error = _http_error(502, b"<html>Bad Gateway</html>")
        with pytest.raises(GraphRequestError, match=r"PATCH failed: HTTP 502$"):
            _client(RaisingConnection(error)).update_contact("graph-1", {})

    def test_http_error_while_unauthenticated_raises_auth_error(self):
        error = _http_error(401, {"error": {"code": "InvalidAuthenticationToken"}})
        with pytest.raises(MSAuthError):
            _client(RaisingConnection(error), authenticated=False).update_contact(
                "graph-1", {}
            )


class TestDeleteContact:
    def test_success(self):
        con = FakeConnection(delete=FakeResponse({}))
        _client(con).delete_contact("graph-1")
        assert con.delete_calls[0].endswith("/me/contacts/graph-1")

    def test_failed_delete_while_authenticated_raises_request_error(self):
        con = FakeConnection(delete=FakeResponse(ok=False))
        with pytest.raises(GraphRequestError, match="deletion"):
            _client(con).delete_contact("graph-1")

    def test_failed_delete_while_unauthenticated_raises_auth_error(self):
        con = FakeConnection(delete=FakeResponse(ok=False))
        with pytest.raises(MSAuthError):
            _client(con, authenticated=False).delete_contact("graph-1")


class TestListContactsRaw:
    def test_paging_combines_pages_and_drops_params_on_next_link(self):
        con = FakeConnection(
            get=[
                FakeResponse(
                    {
                        "value": [{"id": "a"}],
                        "@odata.nextLink": "https://graph.example/page2",
                    }
                ),
                FakeResponse({"value": [{"id": "b"}]}),
            ]
        )
        assert _client(con).list_contacts_raw("id,displayName") == [
            {"id": "a"},
            {"id": "b"},
        ]
        first_url, first_params = con.get_calls[0]
        second_url, second_params = con.get_calls[1]
        assert first_params == {"$top": 100, "$select": "id,displayName"}
        assert second_url == "https://graph.example/page2"
        # nextLink already encodes the query; re-sending $select/$top would 400.
        assert second_params == {}

    def test_first_page_failure_while_authenticated_raises_not_empty_list(self):
        con = FakeConnection(get=[FakeResponse(ok=False)])
        with pytest.raises(GraphRequestError, match="listing"):
            _client(con).list_contacts_raw()

    def test_first_page_failure_while_unauthenticated_raises_auth_error(self):
        con = FakeConnection(get=[FakeResponse(ok=False)])
        with pytest.raises(MSAuthError):
            _client(con, authenticated=False).list_contacts_raw()

    def test_later_page_failure_discards_partial_result(self):
        con = FakeConnection(
            get=[
                FakeResponse(
                    {
                        "value": [{"id": "a"}],
                        "@odata.nextLink": "https://graph.example/page2",
                    }
                ),
                FakeResponse(ok=False),
            ]
        )
        with pytest.raises(GraphRequestError, match="listing"):
            _client(con).list_contacts_raw()

    def test_unparsable_body_raises(self):
        con = FakeConnection(get=[FakeResponse(json_error=ValueError("html page"))])
        with pytest.raises(GraphRequestError, match="unparsable"):
            _client(con).list_contacts_raw()

    def test_non_object_body_raises(self):
        con = FakeConnection(get=[FakeResponse(["not", "a", "dict"])])
        with pytest.raises(GraphRequestError, match="expected a JSON object"):
            _client(con).list_contacts_raw()

    def test_non_list_value_raises(self):
        con = FakeConnection(get=[FakeResponse({"value": {"id": "a"}})])
        with pytest.raises(GraphRequestError, match="non-list"):
            _client(con).list_contacts_raw()

    def test_empty_result_is_a_valid_success(self):
        con = FakeConnection(get=[FakeResponse({"value": []})])
        assert _client(con).list_contacts_raw() == []


class FakeTokenConnection(FakeConnection):
    """A connection that also carries O365's token backend and session."""

    def __init__(self, *, expires_in: float | None, session=True, **kwargs):
        super().__init__(**kwargs)
        self.token_backend = MagicMock()
        self.token_backend.get_access_token.return_value = (
            None
            if expires_in is None
            else {"secret": "at", "expires_on": str(int(time.time() + expires_in))}
        )
        self.session = MagicMock() if session else None
        self.update_session_auth_header = MagicMock()


@pytest.fixture
def refresh():
    with patch.object(graph_client, "refresh_access_token") as mock:
        yield mock


class TestEnsureToken:
    """#38: every run checks the Microsoft sign-in, not only runs with changes."""

    def test_token_valid_for_the_run_is_left_alone(self, refresh):
        con = FakeTokenConnection(expires_in=3600)

        _client(con).ensure_token()

        refresh.assert_not_called()

    @pytest.mark.parametrize("expires_on", ["not-a-number", {"bad": 1}])
    def test_unreadable_cache_requires_reauth(self, refresh, expires_on):
        """Not a generic error on every run: only a sign-in repairs the cache."""
        con = FakeTokenConnection(expires_in=3600)
        con.token_backend.get_access_token.return_value = {"expires_on": expires_on}

        with pytest.raises(MSAuthError):
            _client(con).ensure_token()

        refresh.assert_not_called()

    @pytest.mark.parametrize("expires_in", [-60, 120, None])
    def test_expired_or_expiring_token_is_refreshed(self, refresh, expires_in):
        refresh.return_value = (CredentialStatus.VALID, "new-at")
        con = FakeTokenConnection(expires_in=expires_in)

        _client(con).ensure_token()

        refresh.assert_called_once()
        assert refresh.call_args.args[1] is con.token_backend
        con.update_session_auth_header.assert_called_once_with(access_token="new-at")

    def test_refresh_without_a_session_leaves_the_header_to_o365(self, refresh):
        refresh.return_value = (CredentialStatus.VALID, "new-at")
        con = FakeTokenConnection(expires_in=-60, session=False)

        _client(con).ensure_token()

        con.update_session_auth_header.assert_not_called()

    @pytest.mark.parametrize(
        "status", [CredentialStatus.INVALID_CLIENT, CredentialStatus.REAUTHORIZE]
    )
    def test_rejected_refresh_requires_reauth(self, refresh, status):
        refresh.return_value = (status, None)

        with pytest.raises(MSAuthError):
            _client(FakeTokenConnection(expires_in=-60)).ensure_token()

    def test_missing_refresh_token_requires_reauth(self, refresh):
        con = FakeTokenConnection(expires_in=-60)

        with pytest.raises(MSAuthError):
            _client(con, authenticated=False).ensure_token()

        refresh.assert_not_called()

    def test_unreachable_token_endpoint_is_not_an_auth_failure(self, refresh):
        refresh.return_value = (CredentialStatus.CANNOT_CONNECT, None)

        with pytest.raises(GraphRequestError):
            _client(FakeTokenConnection(expires_in=-60)).ensure_token()

    def test_network_error_is_not_an_auth_failure(self, refresh):
        refresh.side_effect = ConnectionError("dns")

        with pytest.raises(GraphRequestError):
            _client(FakeTokenConnection(expires_in=-60)).ensure_token()


class FailingRefreshConnection(FakeTokenConnection):
    """O365 reports a failed refresh of its own as a bare RuntimeError."""

    def post(self, url, data=None):
        raise RuntimeError("Refresh token operation failed: invalid_client")


class TestRefreshFailureDuringARun:
    """#38: O365's RuntimeError must not pass as a per-contact failure."""

    @pytest.mark.parametrize(
        "status", [CredentialStatus.INVALID_CLIENT, CredentialStatus.REAUTHORIZE]
    )
    def test_revoked_credential_becomes_ms_auth_error(self, refresh, status):
        refresh.return_value = (status, None)
        client = _client(FailingRefreshConnection(expires_in=-60))

        with pytest.raises(MSAuthError):
            client.create_contact({"givenName": "A"})

    def test_transient_failure_stays_a_request_error(self, refresh):
        refresh.return_value = (CredentialStatus.VALID, "new-at")
        client = _client(FailingRefreshConnection(expires_in=-60))

        with pytest.raises(GraphRequestError) as excinfo:
            client.create_contact({"givenName": "A"})

        assert not isinstance(excinfo.value, MSAuthError)
