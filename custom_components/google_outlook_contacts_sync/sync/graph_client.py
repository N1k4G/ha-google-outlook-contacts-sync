"""Microsoft Graph contacts client via python-o365."""

from __future__ import annotations

import logging
from typing import Any

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
    """

    def __init__(self, account: Account) -> None:
        self._account = account

    def _require_auth(self) -> None:
        """Raise MSAuthError when MSAL can no longer provide a valid token."""
        if not self._account.is_authenticated:
            raise MSAuthError(
                "Microsoft token is invalid or expired; re-authentication required."
            )

    def list_contacts(self) -> list[Any]:
        """Return all contacts from the default address book."""
        return list(self._account.address_book().get_contacts(limit=None))

    def list_contacts_raw(self, select: str | None = None) -> list[dict[str, Any]]:
        """Return all contacts as raw Graph JSON dicts, following paging.

        Used for duplicate detection — avoids the O365 Contact wrapper whose
        property names are unreliable.
        """
        params: dict[str, Any] = {"$top": 100}
        if select:
            params["$select"] = select

        contacts: list[dict[str, Any]] = []
        url: str | None = _CONTACTS_URL
        while url:
            response = self._account.con.get(url, params=params)
            if not response:
                self._require_auth()
                break
            body = response.json()
            contacts.extend(body.get("value", []))
            url = body.get("@odata.nextLink")
            # nextLink already encodes query params; don't re-send them.
            params = {}
        return contacts

    def create_contact(self, fields: dict[str, Any]) -> str:
        """Create a new contact and return its Graph contact id."""
        result = self._account.con.post(_CONTACTS_URL, data=fields)
        if not result:
            self._require_auth()
            _LOGGER.error("Failed to create Graph contact")
            return ""
        contact_id = str(result.get("id", ""))
        _LOGGER.debug("Created Graph contact: %s", contact_id)
        return contact_id

    def update_contact(self, contact_id: str, fields: dict[str, Any]) -> None:
        """Update an existing contact in place."""
        result = self._account.con.patch(f"{_CONTACTS_URL}/{contact_id}", data=fields)
        if not result:
            self._require_auth()
        _LOGGER.debug("Updated Graph contact: %s", contact_id)

    def delete_contact(self, contact_id: str) -> None:
        """Delete a contact by its Graph id."""
        result = self._account.con.delete(f"{_CONTACTS_URL}/{contact_id}")
        if not result:
            self._require_auth()
        _LOGGER.debug("Deleted Graph contact: %s", contact_id)


class MSAuthError(Exception):
    """Raised when Microsoft credentials are invalid or expired."""
