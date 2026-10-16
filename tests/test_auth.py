"""Tests for the credential checks behind reconfigure (auth/*.check_credentials).

The provider SDKs are mocked at their boundary: these tests pin how each
provider answer is classified, which decides between saving, a form error and
a fresh sign-in.
"""

from __future__ import annotations

import logging
from unittest.mock import MagicMock, patch

import pytest
from google.auth.exceptions import RefreshError, TransportError

from custom_components.google_outlook_contacts_sync.auth import (
    CredentialStatus,
    google_auth,
    ms_auth,
)

SECRET = "s3cr3t-value"

GOOGLE_TOKEN = {
    "token": "old-access",
    "refresh_token": "grefresh",
    "token_uri": "https://oauth2.googleapis.com/token",
    "client_id": "gid",
    "client_secret": "old-secret",
    "scopes": ["https://www.googleapis.com/auth/contacts.readonly"],
}


def _ms_account(result: object, *, authenticated: bool = True) -> tuple:
    account = MagicMock()
    account.is_authenticated = authenticated
    account.con.msal_client.get_accounts.return_value = [{"username": "u"}]
    account.con.msal_client.acquire_token_silent_with_error.return_value = result
    backend = MagicMock()
    backend.get_token_scopes.return_value = ["https://graph.microsoft.com/User.Read"]
    backend.get_token_dict.return_value = {"refreshed": True}
    return account, backend


class TestMicrosoftCheck:
    def _check(self, account_and_backend: tuple):
        with patch.object(
            ms_auth, "build_account", return_value=account_and_backend
        ) as build:
            result = ms_auth.check_credentials("mid", SECRET, {"cache": 1})
        build.assert_called_once_with("mid", SECRET, {"cache": 1})
        return result

    def test_success_returns_the_refreshed_cache(self) -> None:
        account, backend = _ms_account({"access_token": "at"})

        result = self._check((account, backend))

        assert result.status is CredentialStatus.VALID
        assert result.token == {"refreshed": True}
        backend.save_token.assert_called_once()

    def test_refresh_is_forced(self) -> None:
        """A cached access token would skip presenting the secret entirely."""
        account, backend = _ms_account({"access_token": "at"})

        self._check((account, backend))

        call = account.con.msal_client.acquire_token_silent_with_error.call_args
        assert call.kwargs["force_refresh"] is True

    @pytest.mark.parametrize(
        ("error", "status"),
        [
            ("invalid_client", CredentialStatus.INVALID_CLIENT),
            ("unauthorized_client", CredentialStatus.INVALID_CLIENT),
            ("invalid_grant", CredentialStatus.REAUTHORIZE),
            ("interaction_required", CredentialStatus.REAUTHORIZE),
            ("temporarily_unavailable", CredentialStatus.CANNOT_CONNECT),
        ],
    )
    def test_provider_errors_are_classified(self, error, status) -> None:
        account, backend = _ms_account(
            {"error": error, "error_description": "AADSTS..."}
        )

        result = self._check((account, backend))

        assert result.status is status
        assert result.token is None
        backend.save_token.assert_not_called()

    def test_no_refresh_token_for_this_client(self) -> None:
        account, backend = _ms_account(None)

        assert self._check((account, backend)).status is CredentialStatus.REAUTHORIZE

    def test_unauthenticated_cache_needs_a_sign_in(self) -> None:
        account, backend = _ms_account({"access_token": "at"}, authenticated=False)

        assert self._check((account, backend)).status is CredentialStatus.REAUTHORIZE
        account.con.msal_client.acquire_token_silent_with_error.assert_not_called()

    def test_cache_without_scopes_needs_a_sign_in(self) -> None:
        account, backend = _ms_account({"access_token": "at"})
        backend.get_token_scopes.return_value = None

        assert self._check((account, backend)).status is CredentialStatus.REAUTHORIZE

    def test_network_failure_is_not_a_verdict(self, caplog) -> None:
        account, backend = _ms_account(None)
        account.con.msal_client.acquire_token_silent_with_error.side_effect = (
            ConnectionError(f"boom {SECRET}")
        )

        with caplog.at_level(logging.DEBUG):
            result = self._check((account, backend))

        assert result.status is CredentialStatus.CANNOT_CONNECT
        assert SECRET not in caplog.text


class TestGoogleCheck:
    def test_success_carries_the_new_secret(self) -> None:
        def _refresh(creds, _request):
            creds.token = "new-access"

        with patch.object(google_auth.Credentials, "refresh", _refresh):
            result = google_auth.check_credentials("gid", SECRET, GOOGLE_TOKEN)

        assert result.status is CredentialStatus.VALID
        assert result.token == {
            **GOOGLE_TOKEN,
            "token": "new-access",
            "client_secret": SECRET,
        }

    def test_stored_access_token_is_not_reused(self) -> None:
        """The refresh must reach Google, or the secret goes untested."""
        seen = {}

        def _refresh(creds, _request):
            seen["token"] = creds.token
            seen["secret"] = creds.client_secret

        with patch.object(google_auth.Credentials, "refresh", _refresh):
            google_auth.check_credentials("gid", SECRET, GOOGLE_TOKEN)

        assert seen == {"token": None, "secret": SECRET}

    @pytest.mark.parametrize(
        ("exc", "status"),
        [
            (
                RefreshError(
                    "invalid_client: Unauthorized", {"error": "invalid_client"}
                ),
                CredentialStatus.INVALID_CLIENT,
            ),
            (
                RefreshError("invalid_grant: revoked", {"error": "invalid_grant"}),
                CredentialStatus.REAUTHORIZE,
            ),
            (
                RefreshError("unauthorized_client", {"error": "unauthorized_client"}),
                CredentialStatus.REAUTHORIZE,
            ),
            (RefreshError("no details"), CredentialStatus.REAUTHORIZE),
            (
                RefreshError("server", {"error": "internal"}, retryable=True),
                CredentialStatus.CANNOT_CONNECT,
            ),
            (TransportError("dns"), CredentialStatus.CANNOT_CONNECT),
        ],
    )
    def test_failures_are_classified(self, exc, status, caplog) -> None:
        with (
            patch.object(google_auth.Credentials, "refresh", side_effect=exc),
            caplog.at_level(logging.DEBUG),
        ):
            result = google_auth.check_credentials("gid", SECRET, GOOGLE_TOKEN)

        assert result.status is status
        assert result.token is None
        assert SECRET not in caplog.text

    @pytest.mark.parametrize("token", [None, {}, {**GOOGLE_TOKEN, "refresh_token": ""}])
    def test_missing_refresh_token_needs_a_sign_in(self, token) -> None:
        with patch.object(google_auth.Credentials, "refresh") as refresh:
            result = google_auth.check_credentials("gid", SECRET, token)

        assert result.status is CredentialStatus.REAUTHORIZE
        refresh.assert_not_called()
