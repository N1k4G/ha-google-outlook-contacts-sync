"""Microsoft Graph contacts client via python-o365."""

from __future__ import annotations

import logging
from typing import Any, Never

from O365 import Account  # type: ignore[import-untyped]

_LOGGER = logging.getLogger(__name__)

_CONTACTS_URL = "https://graph.microsoft.com/v1.0/me/contacts"


class GraphContactsClient:
    """Wraps the python-o365 contacts API.

    All public methods are synchronous — call them via
    ``hass.async_add_executor_job``.

    Write operations send the Graph-compatible JSON dict produced by
    ``mapping.to_graph_contact`` directly via the O365 connection object,
    bypassing the O365 Contact wrapper (which lacks birthday, has wrong
    property names for givenName/fileAs, and no get_contact-by-id method).

    Every request failure raises: :class:`MSAuthError` when the account can no
    longer authenticate, otherwise :class:`GraphRequestError`. No method ever
    returns a partial or empty result to signal failure.
    """

    def __init__(self, account: Account) -> None:
        self._account = account

    def _require_auth(self) -> None:
        """Raise MSAuthError when MSAL can no longer provide a valid token."""
        if not self._account.is_authenticated:
            raise MSAuthError(
                "Microsoft token is invalid or expired; re-authentication required."
            )

    def _raise_request_failure(self, operation: str) -> Never:
        """Raise the right error for a failed Graph request.

        Auth problems must reach the coordinator's reauth path, so they are
        classified first; anything else is a plain request failure.
        """
        self._require_auth()
        raise GraphRequestError(f"Microsoft Graph {operation} failed")

    def _response_body(self, response: Any, operation: str) -> dict[str, Any]:
        """Return the JSON object body of a response, or raise GraphRequestError.

        An HTML error page or a truncated body must never be mistaken for an
        empty-but-valid result.
        """
        try:
            body = response.json()
        except Exception as exc:
            raise GraphRequestError(
                f"Microsoft Graph {operation} returned an unparsable body"
            ) from exc
        if not isinstance(body, dict):
            raise GraphRequestError(
                f"Microsoft Graph {operation} returned {type(body).__name__}, "
                "expected a JSON object"
            )
        return body

    def list_contacts(self) -> list[Any]:
        """Return all contacts from the default address book."""
        return list(self._account.address_book().get_contacts(limit=None))

    def list_contacts_raw(self, select: str | None = None) -> list[dict[str, Any]]:
        """Return all contacts as raw Graph JSON dicts, following paging.

        Used for duplicate detection and full-resync reconciliation — avoids the
        O365 Contact wrapper whose property names are unreliable.

        Raises rather than returning a partial list: callers treat the result as
        the complete Outlook state and would otherwise create duplicates or drop
        mappings for contacts that merely failed to load.
        """
        params: dict[str, Any] = {"$top": 100}
        if select:
            params["$select"] = select

        contacts: list[dict[str, Any]] = []
        url: str | None = _CONTACTS_URL
        while url:
            response = self._account.con.get(url, params=params)
            if not response:
                self._raise_request_failure("contact listing")
            body = self._response_body(response, "contact listing")
            value = body.get("value", [])
            if not isinstance(value, list):
                raise GraphRequestError(
                    "Microsoft Graph contact listing returned a non-list 'value'"
                )
            contacts.extend(value)
            url = body.get("@odata.nextLink")
            # nextLink already encodes query params; don't re-send them.
            params = {}
        return contacts

    def create_contact(self, fields: dict[str, Any]) -> str:
        """Create a new contact and return its Graph contact id."""
        response = self._account.con.post(_CONTACTS_URL, data=fields)
        if not response:
            self._raise_request_failure("contact creation")

        # con.post() returns a requests.Response — the id lives in the JSON
        # body, not on the response object itself.
        body = self._response_body(response, "contact creation")
        contact_id = str(body.get("id", ""))
        if not contact_id:
            raise GraphRequestError(
                "Microsoft Graph contact creation returned no contact id"
            )
        _LOGGER.debug("Created Graph contact: %s", contact_id)
        return contact_id

    def update_contact(self, contact_id: str, fields: dict[str, Any]) -> None:
        """Update an existing contact in place."""
        response = self._account.con.patch(f"{_CONTACTS_URL}/{contact_id}", data=fields)
        if not response:
            self._raise_request_failure(f"contact update ({contact_id})")
        _LOGGER.debug("Updated Graph contact: %s", contact_id)

    def delete_contact(self, contact_id: str) -> None:
        """Delete a contact by its Graph id."""
        response = self._account.con.delete(f"{_CONTACTS_URL}/{contact_id}")
        if not response:
            self._raise_request_failure(f"contact deletion ({contact_id})")
        _LOGGER.debug("Deleted Graph contact: %s", contact_id)


class MSAuthError(Exception):
    """Raised when Microsoft credentials are invalid or expired."""


class GraphRequestError(Exception):
    """Raised when a Microsoft Graph request fails while auth is still valid."""
