"""Tests for config_flow.py.

Like test_init.py, the flow runs against a mocked Home Assistant: HA's test
runner imports fcntl and cannot load on Windows. The flow handler's own result
builders (forms, external steps, update-and-abort) are real.
"""

from __future__ import annotations

from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
import voluptuous as vol
from homeassistant.config_entries import SOURCE_REAUTH, SOURCE_RECONFIGURE
from homeassistant.data_entry_flow import FlowResultType

from custom_components.google_outlook_contacts_sync import config_flow
from custom_components.google_outlook_contacts_sync.auth import (
    CredentialCheck,
    CredentialStatus,
)
from custom_components.google_outlook_contacts_sync.config_flow import (
    _OPTIONS_SCHEMA,
    GoogleOutlookContactsSyncConfigFlow,
)
from custom_components.google_outlook_contacts_sync.const import (
    CONF_DELETE_REMOVED,
    CONF_GOOGLE_CLIENT_ID,
    CONF_GOOGLE_CLIENT_SECRET,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    CONF_SYNC_INTERVAL_HOURS,
    DATA_REAUTH_PROVIDER,
    DEFAULT_DELETE_REMOVED,
    DEFAULT_SYNC_INTERVAL_HOURS,
    DOMAIN,
)

MS_TOKEN = {"RefreshToken": {"rt": {"secret": "old-ms-rt"}}}
GOOGLE_TOKEN = {
    "token": "gaccess",
    "refresh_token": "grefresh",
    "token_uri": "https://oauth2.googleapis.com/token",
    "client_id": "gid",
    "client_secret": "gsecret",
    "scopes": ["https://www.googleapis.com/auth/contacts.readonly"],
}


class TestOptionsFlow:
    def test_options_schema_default_values(self) -> None:
        result = _OPTIONS_SCHEMA(
            {
                CONF_SYNC_INTERVAL_HOURS: DEFAULT_SYNC_INTERVAL_HOURS,
                CONF_DELETE_REMOVED: DEFAULT_DELETE_REMOVED,
            }
        )
        assert result[CONF_SYNC_INTERVAL_HOURS] == DEFAULT_SYNC_INTERVAL_HOURS
        assert result[CONF_DELETE_REMOVED] == DEFAULT_DELETE_REMOVED

    def test_options_schema_rejects_zero_interval(self) -> None:
        with pytest.raises(vol.Invalid):
            _OPTIONS_SCHEMA({CONF_SYNC_INTERVAL_HOURS: 0, CONF_DELETE_REMOVED: False})

    def test_options_schema_rejects_interval_above_max(self) -> None:
        with pytest.raises(vol.Invalid):
            _OPTIONS_SCHEMA({CONF_SYNC_INTERVAL_HOURS: 200, CONF_DELETE_REMOVED: False})


def _make_entry(**extra: Any) -> MagicMock:
    entry = MagicMock()
    entry.entry_id = "entry-1"
    entry.title = "Google → Outlook Contacts"
    entry.update_listeners = []
    entry.data = {
        CONF_MS_CLIENT_ID: "mid",
        CONF_MS_CLIENT_SECRET: "msecret",
        "ms_token": MS_TOKEN,
        CONF_GOOGLE_CLIENT_ID: "gid",
        CONF_GOOGLE_CLIENT_SECRET: "gsecret",
        "google_token": GOOGLE_TOKEN,
        **extra,
    }
    return entry


def _make_flow(source: str, entry: MagicMock) -> GoogleOutlookContactsSyncConfigFlow:
    hass = MagicMock()
    hass.data = {}
    hass.config.external_url = "https://ha.example"
    hass.async_add_executor_job = AsyncMock(side_effect=lambda fn, *a: fn(*a))
    hass.config_entries.async_get_entry = MagicMock(return_value=entry)
    hass.config_entries.async_get_known_entry = MagicMock(return_value=entry)
    hass.config_entries.async_update_entry = MagicMock(return_value=True)
    hass.config_entries.async_schedule_reload = MagicMock()

    flow = GoogleOutlookContactsSyncConfigFlow()
    flow.hass = hass
    flow.handler = DOMAIN
    flow.flow_id = "flow-1"
    flow.context = {"source": source, "entry_id": entry.entry_id}
    return flow


def _written_data(flow: GoogleOutlookContactsSyncConfigFlow) -> dict[str, Any]:
    update = flow.hass.config_entries.async_update_entry
    update.assert_called_once()
    return dict(update.call_args.kwargs["data"])


def _reconfigure_input(**overrides: str) -> dict[str, str]:
    return {CONF_MS_CLIENT_ID: "mid", CONF_GOOGLE_CLIENT_ID: "gid", **overrides}


@pytest.fixture
def ms_check():
    with patch.object(config_flow, "ms_check_credentials") as mock:
        yield mock


@pytest.fixture
def google_check():
    with patch.object(config_flow, "google_check_credentials") as mock:
        yield mock


@pytest.fixture
def ms_oauth():
    """Stand-ins for the Microsoft OAuth plumbing used by the external step."""
    backend = MagicMock()
    backend.get_token_dict.return_value = {"fresh": "ms-token"}
    with (
        patch.object(config_flow, "async_register_callback_views"),
        patch.object(
            config_flow, "build_account", return_value=(MagicMock(), backend)
        ) as build_account,
        patch.object(
            config_flow,
            "ms_get_auth_url",
            return_value=("https://login.example/ms", {"state": "ms-state"}),
        ),
        patch.object(config_flow, "request_token", return_value=True),
    ):
        yield build_account


@pytest.fixture
def google_oauth():
    """Stand-ins for the Google OAuth plumbing used by the external step."""
    with (
        patch.object(config_flow, "async_register_callback_views"),
        patch.object(config_flow, "build_auth_flow") as build_auth_flow,
        patch.object(
            config_flow,
            "get_auth_url",
            return_value=("https://accounts.example/google", "g-state"),
        ),
        patch.object(
            config_flow, "exchange_code", return_value={"fresh": "google-token"}
        ),
    ):
        yield build_auth_flow


class TestReconfigure:
    async def test_form_prefills_ids_but_never_secrets(self) -> None:
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure()

        assert result["type"] is FlowResultType.FORM
        assert result["step_id"] == "reconfigure"
        suggested = {
            str(key): (key.description or {}).get("suggested_value")
            for key in result["data_schema"].schema
        }
        assert suggested[CONF_MS_CLIENT_ID] == "mid"
        assert suggested[CONF_GOOGLE_CLIENT_ID] == "gid"
        assert suggested[CONF_MS_CLIENT_SECRET] is None
        assert suggested[CONF_GOOGLE_CLIENT_SECRET] is None

    async def test_new_ms_secret_keeps_tokens_and_skips_oauth(
        self, ms_check, google_check
    ) -> None:
        ms_check.return_value = CredentialCheck(
            CredentialStatus.VALID, {"refreshed": "ms-token"}
        )
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_MS_CLIENT_SECRET: " new-msecret "})
        )

        assert result["type"] is FlowResultType.ABORT
        assert result["reason"] == "reconfigure_successful"
        ms_check.assert_called_once_with("mid", "new-msecret", MS_TOKEN)
        google_check.assert_not_called()
        data = _written_data(flow)
        assert data[CONF_MS_CLIENT_SECRET] == "new-msecret"
        assert data["ms_token"] == {"refreshed": "ms-token"}
        assert data[CONF_GOOGLE_CLIENT_SECRET] == "gsecret"
        assert data["google_token"] == GOOGLE_TOKEN
        flow.hass.config_entries.async_schedule_reload.assert_called_once_with(
            "entry-1"
        )

    async def test_new_google_secret_is_stored_in_the_token_too(
        self, ms_check, google_check
    ) -> None:
        """Setup rebuilds the People client from google_token, not entry keys."""
        refreshed = {**GOOGLE_TOKEN, "client_secret": "new-gsecret", "token": "new"}
        google_check.return_value = CredentialCheck(CredentialStatus.VALID, refreshed)
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_GOOGLE_CLIENT_SECRET: "new-gsecret"})
        )

        assert result["type"] is FlowResultType.ABORT
        ms_check.assert_not_called()
        google_check.assert_called_once_with("gid", "new-gsecret", GOOGLE_TOKEN)
        data = _written_data(flow)
        assert data[CONF_GOOGLE_CLIENT_SECRET] == "new-gsecret"
        assert data["google_token"] == refreshed
        assert data["ms_token"] == MS_TOKEN

    async def test_empty_secret_fields_keep_the_stored_secrets(
        self, ms_check, google_check
    ) -> None:
        entry = _make_entry()
        flow = _make_flow(SOURCE_RECONFIGURE, entry)

        result = await flow.async_step_reconfigure(
            _reconfigure_input(
                **{CONF_MS_CLIENT_SECRET: "", CONF_GOOGLE_CLIENT_SECRET: "  "}
            )
        )

        assert result["type"] is FlowResultType.ABORT
        ms_check.assert_not_called()
        google_check.assert_not_called()
        assert _written_data(flow) == entry.data

    @pytest.mark.parametrize(
        ("provider", "status", "error"),
        [
            ("ms", CredentialStatus.INVALID_CLIENT, "ms_invalid_client"),
            ("google", CredentialStatus.INVALID_CLIENT, "google_invalid_client"),
            ("ms", CredentialStatus.CANNOT_CONNECT, "cannot_connect"),
            ("google", CredentialStatus.CANNOT_CONNECT, "cannot_connect"),
        ],
    )
    async def test_rejected_secret_saves_nothing(
        self, ms_check, google_check, provider, status, error
    ) -> None:
        ms_check.return_value = CredentialCheck(CredentialStatus.VALID, {"x": 1})
        google_check.return_value = CredentialCheck(CredentialStatus.VALID, {"x": 1})
        (ms_check if provider == "ms" else google_check).return_value = CredentialCheck(
            status
        )
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(
                **{
                    CONF_MS_CLIENT_SECRET: "new-msecret",
                    CONF_GOOGLE_CLIENT_SECRET: "new-gsecret",
                }
            )
        )

        assert result["type"] is FlowResultType.FORM
        assert result["step_id"] == "reconfigure"
        assert result["errors"] == {"base": error}
        flow.hass.config_entries.async_update_entry.assert_not_called()
        flow.hass.config_entries.async_schedule_reload.assert_not_called()

    async def test_a_failed_attempt_does_not_leak_into_the_next(
        self, ms_check, google_check
    ) -> None:
        """Staged credentials are reset from the entry on every submit."""
        ms_check.return_value = CredentialCheck(CredentialStatus.VALID, {"x": 1})
        google_check.return_value = CredentialCheck(CredentialStatus.INVALID_CLIENT)
        entry = _make_entry()
        flow = _make_flow(SOURCE_RECONFIGURE, entry)
        await flow.async_step_reconfigure(
            _reconfigure_input(
                **{
                    CONF_MS_CLIENT_SECRET: "new-msecret",
                    CONF_GOOGLE_CLIENT_SECRET: "bad-gsecret",
                }
            )
        )

        result = await flow.async_step_reconfigure(_reconfigure_input())

        assert result["type"] is FlowResultType.ABORT
        assert _written_data(flow) == entry.data

    async def test_missing_client_id_is_rejected(self, ms_check) -> None:
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_MS_CLIENT_ID: "  "})
        )

        assert result["errors"] == {"base": "missing_client_id"}
        ms_check.assert_not_called()
        flow.hass.config_entries.async_update_entry.assert_not_called()

    @pytest.mark.parametrize("field", [CONF_MS_CLIENT_ID, CONF_GOOGLE_CLIENT_ID])
    async def test_new_client_id_requires_its_secret(self, field) -> None:
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{field: "other-client"})
        )

        assert result["errors"] == {"base": "new_client_requires_secret"}
        flow.hass.config_entries.async_update_entry.assert_not_called()

    async def test_new_ms_client_id_signs_in_again(
        self, ms_check, google_check, ms_oauth
    ) -> None:
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(
                **{CONF_MS_CLIENT_ID: "mid-2", CONF_MS_CLIENT_SECRET: "msecret-2"}
            )
        )

        # A refresh token is bound to its client ID; no point trying it.
        ms_check.assert_not_called()
        assert result["type"] is FlowResultType.EXTERNAL_STEP
        assert result["step_id"] == "ms_auth_url"
        ms_oauth.assert_called_once_with("mid-2", "msecret-2")
        flow.hass.config_entries.async_update_entry.assert_not_called()

        await flow.async_step_ms_auth_url({"authorization_url": "https://cb"})
        result = await flow.async_step_ms_auth_callback()

        assert result["type"] is FlowResultType.ABORT
        assert result["reason"] == "reconfigure_successful"
        data = _written_data(flow)
        assert data[CONF_MS_CLIENT_ID] == "mid-2"
        assert data[CONF_MS_CLIENT_SECRET] == "msecret-2"
        assert data["ms_token"] == {"fresh": "ms-token"}
        assert data["google_token"] == GOOGLE_TOKEN

    async def test_unusable_ms_token_signs_in_again(self, ms_check, ms_oauth) -> None:
        ms_check.return_value = CredentialCheck(CredentialStatus.REAUTHORIZE)
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_MS_CLIENT_SECRET: "new-msecret"})
        )

        assert result["type"] is FlowResultType.EXTERNAL_STEP
        assert result["step_id"] == "ms_auth_url"
        ms_oauth.assert_called_once_with("mid", "new-msecret")

    async def test_unusable_google_token_signs_in_again(
        self, google_check, google_oauth
    ) -> None:
        google_check.return_value = CredentialCheck(CredentialStatus.REAUTHORIZE)
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_GOOGLE_CLIENT_SECRET: "new-gsecret"})
        )

        assert result["type"] is FlowResultType.EXTERNAL_STEP
        assert result["step_id"] == "google_auth_url"
        assert google_oauth.call_args.args[:2] == ("gid", "new-gsecret")

        await flow.async_step_google_auth_url({"code": "auth-code"})
        result = await flow.async_step_google_auth_callback()

        assert result["reason"] == "reconfigure_successful"
        data = _written_data(flow)
        assert data[CONF_GOOGLE_CLIENT_SECRET] == "new-gsecret"
        assert data["google_token"] == {"fresh": "google-token"}
        assert data["ms_token"] == MS_TOKEN

    async def test_both_providers_sign_in_in_turn(
        self, google_check, ms_oauth, google_oauth
    ) -> None:
        google_check.return_value = CredentialCheck(CredentialStatus.REAUTHORIZE)
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())

        result = await flow.async_step_reconfigure(
            _reconfigure_input(
                **{
                    CONF_MS_CLIENT_ID: "mid-2",
                    CONF_MS_CLIENT_SECRET: "msecret-2",
                    CONF_GOOGLE_CLIENT_SECRET: "new-gsecret",
                }
            )
        )
        assert result["step_id"] == "ms_auth_url"

        await flow.async_step_ms_auth_url({"authorization_url": "https://cb"})
        result = await flow.async_step_ms_auth_callback()
        assert result["type"] is FlowResultType.EXTERNAL_STEP
        assert result["step_id"] == "google_auth_url"
        flow.hass.config_entries.async_update_entry.assert_not_called()

        await flow.async_step_google_auth_url({"code": "auth-code"})
        result = await flow.async_step_google_auth_callback()

        assert result["reason"] == "reconfigure_successful"
        data = _written_data(flow)
        assert data["ms_token"] == {"fresh": "ms-token"}
        assert data["google_token"] == {"fresh": "google-token"}

    async def test_token_refreshed_during_sign_in_is_not_overwritten(
        self, google_check, google_oauth
    ) -> None:
        google_check.return_value = CredentialCheck(CredentialStatus.REAUTHORIZE)
        entry = _make_entry()
        flow = _make_flow(SOURCE_RECONFIGURE, entry)
        await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_GOOGLE_CLIENT_SECRET: "new-gsecret"})
        )

        # The running integration persists a refreshed Microsoft token while
        # the user is still signing in with Google.
        entry.data = {**entry.data, "ms_token": {"runtime": "refreshed"}}
        await flow.async_step_google_auth_url({"code": "auth-code"})
        await flow.async_step_google_auth_callback()

        data = _written_data(flow)
        assert data["ms_token"] == {"runtime": "refreshed"}
        assert data["google_token"] == {"fresh": "google-token"}

    async def test_failed_sign_in_retry_keeps_the_client_id(
        self, ms_check, ms_oauth
    ) -> None:
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry())
        await flow.async_step_reconfigure(
            _reconfigure_input(
                **{CONF_MS_CLIENT_ID: "mid-2", CONF_MS_CLIENT_SECRET: "msecret-2"}
            )
        )

        # The provider redirects back with an error (e.g. consent denied).
        await flow.async_step_ms_auth_url({"error": "access_denied"})
        result = await flow.async_step_ms_auth_callback()

        assert result["type"] is FlowResultType.FORM
        assert result["step_id"] == "ms_credentials"
        assert result["errors"] == {"base": "ms_auth_failed"}
        suggested = {
            str(key): (key.description or {}).get("suggested_value")
            for key in result["data_schema"].schema
        }
        assert suggested[CONF_MS_CLIENT_ID] == "mid-2"
        assert suggested[CONF_MS_CLIENT_SECRET] is None
        flow.hass.config_entries.async_update_entry.assert_not_called()

    async def test_keys_the_flow_does_not_own_survive(self, ms_check) -> None:
        ms_check.return_value = CredentialCheck(CredentialStatus.VALID, {"x": 1})
        flow = _make_flow(SOURCE_RECONFIGURE, _make_entry(delete_removed=True))

        await flow.async_step_reconfigure(
            _reconfigure_input(**{CONF_MS_CLIENT_SECRET: "new-msecret"})
        )

        assert _written_data(flow)[CONF_DELETE_REMOVED] is True


class TestReauth:
    async def test_ms_reauth_offers_a_new_secret(self, ms_oauth) -> None:
        flow = _make_flow(SOURCE_REAUTH, _make_entry())
        flow.hass.data[DOMAIN] = {DATA_REAUTH_PROVIDER: {"entry-1": "ms"}}

        result = await flow.async_step_reauth({})

        assert result["type"] is FlowResultType.FORM
        assert result["step_id"] == "reauth_ms_confirm"
        assert CONF_MS_CLIENT_SECRET in {str(k) for k in result["data_schema"].schema}

        result = await flow.async_step_reauth_ms_confirm(
            {CONF_MS_CLIENT_SECRET: "new-msecret"}
        )
        assert result["step_id"] == "ms_auth_url"
        ms_oauth.assert_called_once_with("mid", "new-msecret")

        await flow.async_step_ms_auth_url({"authorization_url": "https://cb"})
        result = await flow.async_step_ms_auth_callback()

        assert result["reason"] == "reauth_successful"
        data = _written_data(flow)
        assert data[CONF_MS_CLIENT_SECRET] == "new-msecret"
        assert data["ms_token"] == {"fresh": "ms-token"}
        assert data["google_token"] == GOOGLE_TOKEN

    async def test_ms_reauth_without_a_secret_keeps_the_stored_one(
        self, ms_oauth
    ) -> None:
        flow = _make_flow(SOURCE_REAUTH, _make_entry())
        flow.hass.data[DOMAIN] = {DATA_REAUTH_PROVIDER: {"entry-1": "ms"}}
        await flow.async_step_reauth({})

        await flow.async_step_reauth_ms_confirm({})

        ms_oauth.assert_called_once_with("mid", "msecret")

    async def test_google_reauth_offers_a_new_secret(self, google_oauth) -> None:
        flow = _make_flow(SOURCE_REAUTH, _make_entry())

        result = await flow.async_step_reauth({})

        assert result["step_id"] == "reauth_confirm"
        assert CONF_GOOGLE_CLIENT_SECRET in {
            str(k) for k in result["data_schema"].schema
        }

        result = await flow.async_step_reauth_confirm(
            {CONF_GOOGLE_CLIENT_SECRET: "new-gsecret"}
        )
        assert result["step_id"] == "google_auth_url"
        assert google_oauth.call_args.args[:2] == ("gid", "new-gsecret")

        await flow.async_step_google_auth_url({"code": "auth-code"})
        result = await flow.async_step_google_auth_callback()

        assert result["reason"] == "reauth_successful"
        data = _written_data(flow)
        assert data[CONF_GOOGLE_CLIENT_SECRET] == "new-gsecret"
        assert data["google_token"] == {"fresh": "google-token"}
        assert data["ms_token"] == MS_TOKEN

    async def test_google_reauth_without_a_secret_keeps_the_stored_one(
        self, google_oauth
    ) -> None:
        flow = _make_flow(SOURCE_REAUTH, _make_entry())
        await flow.async_step_reauth({})

        await flow.async_step_reauth_confirm({CONF_GOOGLE_CLIENT_SECRET: ""})

        assert google_oauth.call_args.args[:2] == ("gid", "gsecret")
