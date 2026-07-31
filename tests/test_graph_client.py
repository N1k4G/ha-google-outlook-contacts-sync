"""Tests for sync/graph_client.py — the adapter boundary to O365/requests.

These tests deliberately avoid ``MagicMock`` for API responses. ``O365.Connection``
returns a ``requests.Response``, which offers ``.json()`` but *not* ``.get()``.
A MagicMock would answer both and silently hide exactly the class of bug this
module had (calling ``.get("id")`` on the response object).
"""

from __future__ import annotations

import pytest

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
