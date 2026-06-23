"""Google People API client wrapper."""

from __future__ import annotations

import logging
from typing import Any

from google.auth.exceptions import RefreshError  # type: ignore[import-untyped]
from google.oauth2.credentials import Credentials  # type: ignore[import-untyped]
from googleapiclient.discovery import build  # type: ignore[import-untyped]
from googleapiclient.errors import HttpError  # type: ignore[import-untyped]

from ..const import GOOGLE_API_SERVICE, GOOGLE_API_VERSION, GOOGLE_SCOPE

_LOGGER = logging.getLogger(__name__)

_PERSON_FIELDS = (
    "names,emailAddresses,phoneNumbers,birthdays,addresses,organizations,metadata"
)

_PAGE_SIZE = 200


class GoogleContactsClient:
    """Thin async-friendly wrapper around the People API.

    All blocking SDK calls are meant to be executed via
    ``hass.async_add_executor_job``.
    """

    def __init__(self, credentials: Credentials) -> None:
        self._credentials = credentials
        self._service: Any | None = None

    def _get_service(self) -> Any:
        if self._service is None:
            self._service = build(
                GOOGLE_API_SERVICE,
                GOOGLE_API_VERSION,
                credentials=self._credentials,
                cache_discovery=False,
            )
        return self._service

    def fetch_contacts(
        self, sync_token: str | None = None
    ) -> tuple[list[dict[str, Any]], str | None]:
        """Fetch contacts (or delta) from the People API.

        Returns a tuple of (people, next_sync_token).  When *sync_token* is
        provided, only changed/deleted persons since that token are returned.
        """
        service = self._get_service()
        people_resource = service.people().connections()

        all_people: list[dict[str, Any]] = []
        next_sync_token: str | None = None
        page_token: str | None = None

        while True:
            kwargs: dict[str, Any] = {
                "resourceName": "people/me",
                "personFields": _PERSON_FIELDS,
                "pageSize": _PAGE_SIZE,
                "requestSyncToken": True,
            }
            if sync_token:
                kwargs["syncToken"] = sync_token
            if page_token:
                kwargs["pageToken"] = page_token

            try:
                response: dict[str, Any] = people_resource.list(**kwargs).execute()
            except HttpError as exc:
                if exc.resp.status == 410:
                    raise SyncTokenExpiredError(
                        "Google sync token has expired; full re-sync required."
                    ) from exc
                raise
            except RefreshError as exc:
                raise GoogleAuthError(str(exc)) from exc

            connections: list[dict[str, Any]] = response.get("connections", [])
            all_people.extend(connections)

            next_sync_token = response.get("nextSyncToken")
            page_token = response.get("nextPageToken")
            if not page_token:
                break

        _LOGGER.debug(
            "Fetched %d contacts from Google People API (delta=%s)",
            len(all_people),
            sync_token is not None,
        )
        return all_people, next_sync_token

    @classmethod
    def from_token_dict(
        cls, token_dict: dict[str, Any]
    ) -> GoogleContactsClient:
        """Build a client from a serialized token dict."""
        creds = Credentials(  # type: ignore[no-untyped-call]
            token=token_dict.get("token"),
            refresh_token=token_dict.get("refresh_token"),
            token_uri=token_dict.get(
                "token_uri", "https://oauth2.googleapis.com/token"
            ),
            client_id=token_dict["client_id"],
            client_secret=token_dict["client_secret"],
            scopes=[GOOGLE_SCOPE],
        )
        return cls(creds)

    def get_token_dict(self) -> dict[str, Any]:
        """Serialize the current credentials to a storable dict."""
        return {
            "token": self._credentials.token,
            "refresh_token": self._credentials.refresh_token,
            "token_uri": self._credentials.token_uri,
            "client_id": self._credentials.client_id,
            "client_secret": self._credentials.client_secret,
            "scopes": list(self._credentials.scopes or []),
        }


class SyncTokenExpiredError(Exception):
    """Raised when the Google sync token is no longer valid."""


class GoogleAuthError(Exception):
    """Raised when Google credentials are invalid or expired (e.g. invalid_grant)."""
