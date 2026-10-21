"""Microsoft Graph contacts client via python-o365."""

from __future__ import annotations

import logging
import time
from typing import Any, Never

from O365 import Account  # type: ignore[import-untyped]
from requests import HTTPError

from ..auth import CredentialStatus
from ..auth.ms_auth import refresh_access_token

_LOGGER = logging.getLogger(__name__)

_CONTACTS_URL = "https://graph.microsoft.com/v1.0/me/contacts"

# Refresh ahead of expiry so the access token outlasts the run that starts now.
_REFRESH_MARGIN_SECONDS = 300


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

    def ensure_token(self) -> None:
        """Refresh the access token now unless it outlasts the coming run.

        O365 only refreshes after Graph answers 401, so a run that changes
        nothing in Outlook never redeems the refresh token at all: a revoked
        or expired client secret went unnoticed, and the refresh token itself
        lapsed after Entra's inactivity window. Called at the start of every
        run, this surfaces both as MSAuthError and keeps the refresh token in
        use.
        """
        if not self._account.is_authenticated:
            # Also loads the token cache, so the expiry below is the real one.
            raise MSAuthError(
                "Microsoft refresh token is missing; re-authentication required."
            )
        try:
            token = self._account.con.token_backend.get_access_token()
            expires_on = int((token or {}).get("expires_on") or 0)
        except (KeyError, TypeError, ValueError) as exc:
            # A corrupted cache: MSAL's own search trips over the entry, and a
            # refresh may add a new one beside it rather than replace it. Only
            # a fresh sign-in writes a clean cache; without this, every run
            # would fail the same way and never ask for one.
            raise MSAuthError(
                "Microsoft token cache is unreadable; re-authentication required."
            ) from exc
        if expires_on - time.time() > _REFRESH_MARGIN_SECONDS:
            return
        self._refresh()

    def _refresh(self) -> None:
        """Redeem the refresh token; raise MSAuthError if Entra refuses."""
        con = self._account.con
        try:
            status, access_token = refresh_access_token(
                self._account, con.token_backend
            )
        except Exception as exc:
            raise GraphRequestError(
                f"Microsoft token refresh failed: {type(exc).__name__}"
            ) from exc
        if status is CredentialStatus.VALID:
            if con.session is not None and access_token:
                # The session caches the bearer header; point it at the new
                # token. A session created later reads it from the backend.
                con.update_session_auth_header(access_token=access_token)
            return
        if status is CredentialStatus.CANNOT_CONNECT:
            raise GraphRequestError("Microsoft token endpoint is unavailable")
        raise MSAuthError(
            f"Microsoft token refresh rejected ({status}); re-authentication required."
        )

    def _request(self, method: str, url: str, **kwargs: Any) -> Any:
        """Send a Graph request, classifying a failed token refresh.

        When O365's own refresh fails it raises a bare RuntimeError that says
        nothing usable about why. Refreshing again through MSAL tells a
        revoked credential (MSAuthError, so reauth starts) from an outage
        (GraphRequestError, so the run fails and is retried).
        """
        try:
            return getattr(self._account.con, method)(url, **kwargs)
        except RuntimeError as exc:
            self._refresh()
            raise GraphRequestError(
                f"Microsoft Graph {method.upper()} failed during a token refresh"
            ) from exc
        except HTTPError as exc:
            # O365 raises on every 4xx/5xx. Its message is the URL plus Graph's
            # human text; the status and Graph's error code are what tell a
            # rejected payload from an outage, so they lead the message.
            self._require_auth()
            raise GraphRequestError(
                f"Microsoft Graph {method.upper()} failed: {_describe_http_error(exc)}"
            ) from exc

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
            response = self._request("get", url, params=params)
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
        response = self._request("post", _CONTACTS_URL, data=fields)
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
        response = self._request("patch", f"{_CONTACTS_URL}/{contact_id}", data=fields)
        if not response:
            self._raise_request_failure(f"contact update ({contact_id})")
        _LOGGER.debug("Updated Graph contact: %s", contact_id)

    def delete_contact(self, contact_id: str) -> None:
        """Delete a contact by its Graph id."""
        response = self._request("delete", f"{_CONTACTS_URL}/{contact_id}")
        if not response:
            self._raise_request_failure(f"contact deletion ({contact_id})")
        _LOGGER.debug("Deleted Graph contact: %s", contact_id)


def _describe_http_error(exc: HTTPError) -> str:
    """Summarise a Graph error response as "HTTP <status> <code>: <message>"."""
    response = exc.response
    if response is None:
        return str(exc)
    try:
        error = response.json().get("error") or {}
    except Exception:
        # An unparsable body still has a status worth reporting.
        error = {}
    if not isinstance(error, dict):
        error = {}
    code = str(error.get("code") or "")
    message = str(error.get("message") or "")
    parts = [f"HTTP {response.status_code}"]
    if code:
        parts.append(code)
    summary = " ".join(parts)
    return f"{summary}: {message}" if message else summary


class MSAuthError(Exception):
    """Raised when Microsoft credentials are invalid or expired."""


class GraphRequestError(Exception):
    """Raised when a Microsoft Graph request fails while auth is still valid."""
