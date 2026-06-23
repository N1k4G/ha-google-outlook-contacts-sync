"""Config flow and options flow for Google Outlook Contacts Sync."""

from __future__ import annotations

import logging
from typing import Any

import voluptuous as vol
from homeassistant.config_entries import (
    ConfigEntry,
    ConfigFlow,
    ConfigFlowResult,
    OptionsFlow,
)
from homeassistant.core import callback

from .auth.google_auth import build_auth_flow, exchange_code, get_auth_url
from .auth.ms_auth import (
    build_account,
    build_redirect_uri,
    request_token,
)
from .auth.ms_auth import (
    get_auth_url as ms_get_auth_url,
)
from .const import (
    CONF_AUTO_REMOVE_DUPLICATES,
    CONF_DELETE_REMOVED,
    CONF_GOOGLE_CLIENT_ID,
    CONF_GOOGLE_CLIENT_SECRET,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    CONF_SYNC_INTERVAL_HOURS,
    DATA_GOOGLE_OAUTH_FLOWS,
    DATA_MS_OAUTH_FLOWS,
    DEFAULT_AUTO_REMOVE_DUPLICATES,
    DEFAULT_DELETE_REMOVED,
    DEFAULT_SYNC_INTERVAL_HOURS,
    DOMAIN,
    GOOGLE_AUTH_CALLBACK_PATH,
)
from .http import GoogleAuthCallbackView, MSAuthCallbackView

_LOGGER = logging.getLogger(__name__)

_STEP_MS_CREDENTIALS = vol.Schema(
    {
        vol.Required(CONF_MS_CLIENT_ID): str,
        vol.Required(CONF_MS_CLIENT_SECRET): str,
    }
)

_STEP_GOOGLE_CREDENTIALS = vol.Schema(
    {
        vol.Required(CONF_GOOGLE_CLIENT_ID): str,
        vol.Required(CONF_GOOGLE_CLIENT_SECRET): str,
    }
)

_OPTIONS_SCHEMA = vol.Schema(
    {
        vol.Required(
            CONF_SYNC_INTERVAL_HOURS,
            default=DEFAULT_SYNC_INTERVAL_HOURS,
        ): vol.All(int, vol.Range(min=1, max=168)),
        vol.Required(CONF_DELETE_REMOVED, default=DEFAULT_DELETE_REMOVED): bool,
        vol.Required(
            CONF_AUTO_REMOVE_DUPLICATES,
            default=DEFAULT_AUTO_REMOVE_DUPLICATES,
        ): bool,
    }
)


class GoogleOutlookContactsSyncConfigFlow(ConfigFlow, domain=DOMAIN):
    """Config flow for Microsoft and Google OAuth setup."""

    VERSION = 1

    def __init__(self) -> None:
        self._ms_client_id: str = ""
        self._ms_client_secret: str = ""
        self._ms_account: Any | None = None
        self._ms_backend: Any | None = None
        self._ms_token: dict[str, Any] | None = None
        self._ms_auth_flow: dict[str, Any] | None = None
        self._ms_authorization_url: str = ""
        self._google_client_id: str = ""
        self._google_client_secret: str = ""
        self._google_token: dict[str, Any] | None = None
        self._google_flow: Any | None = None
        self._google_code: str = ""
        self._reauth_entry: ConfigEntry | None = None

    async def async_step_user(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 1: enter Microsoft app credentials."""
        if self._async_current_entries():
            return self.async_abort(reason="already_configured")

        return await self.async_step_ms_credentials(user_input)

    async def async_step_reauth(
        self, entry_data: dict[str, Any]
    ) -> ConfigFlowResult:
        """Re-authorize Google when the stored token is invalid or expired."""
        entry_id = self.context.get("entry_id", "")
        reauth_entry = self.hass.config_entries.async_get_entry(entry_id)
        if reauth_entry is None:
            return await self.async_step_user()
        self._reauth_entry = reauth_entry
        config_data = reauth_entry.data
        self._ms_client_id = config_data.get(CONF_MS_CLIENT_ID, "")
        self._ms_client_secret = config_data.get(CONF_MS_CLIENT_SECRET, "")
        self._ms_token = dict(config_data.get("ms_token") or {})
        self._google_client_id = config_data.get(CONF_GOOGLE_CLIENT_ID, "")
        self._google_client_secret = config_data.get(CONF_GOOGLE_CLIENT_SECRET, "")
        return await self.async_step_reauth_confirm()

    async def async_step_reauth_confirm(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Confirm before launching the Google re-authorization flow."""
        if user_input is not None:
            return await self.async_step_google_auth_url()
        return self.async_show_form(
            step_id="reauth_confirm",
            data_schema=vol.Schema({}),
        )

    async def async_step_ms_credentials(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        errors: dict[str, str] = {}

        if user_input is not None:
            self._ms_client_id = user_input[CONF_MS_CLIENT_ID].strip()
            self._ms_client_secret = user_input[CONF_MS_CLIENT_SECRET].strip()
            return await self.async_step_ms_auth_url()

        return self.async_show_form(
            step_id="ms_credentials",
            data_schema=_STEP_MS_CREDENTIALS,
            errors=errors,
        )

    async def async_step_ms_auth_url(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 2: generate MS auth URL and wait for the callback.

        HA requires that an external step only transitions to
        async_external_step_done — not to a form. Token exchange therefore
        happens in async_step_ms_auth_callback.
        """
        if user_input is not None:
            if "error" in user_input:
                _LOGGER.error("MS OAuth error from callback: %s", user_input["error"])
                self._ms_authorization_url = ""
            else:
                self._ms_authorization_url = user_input.get("authorization_url", "")
            return self.async_external_step_done(next_step_id="ms_auth_callback")

        if not self.hass.config.external_url:
            return self.async_abort(reason="missing_external_url")

        # async_setup is not called during first-time setup (no config entry
        # exists yet), so the callback view must be registered here.
        self.hass.http.register_view(MSAuthCallbackView(self.hass))

        try:
            redirect_uri = self._build_redirect_uri()
            self._ms_account, self._ms_backend = build_account(
                self._ms_client_id, self._ms_client_secret
            )
            auth_url, self._ms_auth_flow = await self.hass.async_add_executor_job(
                ms_get_auth_url,
                self._ms_account,
                redirect_uri,
            )
        except Exception:
            _LOGGER.exception("Failed to build Microsoft OAuth URL")
            return self.async_show_form(
                step_id="ms_credentials",
                data_schema=_STEP_MS_CREDENTIALS,
                errors={"base": "ms_url_failed"},
            )

        msal_state: str = (
            self._ms_auth_flow.get("state", "") if self._ms_auth_flow else ""
        )
        self.hass.data.setdefault(DOMAIN, {}).setdefault(
            DATA_MS_OAUTH_FLOWS, {}
        )[msal_state] = self.flow_id

        return self.async_external_step(
            step_id="ms_auth_url",
            url=auth_url,
        )

    async def async_step_ms_auth_callback(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 2b: exchange the MS authorization code for tokens."""
        if (
            not self._ms_authorization_url
            or self._ms_account is None
            or self._ms_backend is None
            or self._ms_auth_flow is None
        ):
            return self.async_show_form(
                step_id="ms_credentials",
                data_schema=_STEP_MS_CREDENTIALS,
                errors={"base": "ms_auth_failed"},
            )

        success = await self.hass.async_add_executor_job(
            request_token,
            self._ms_account,
            self._ms_authorization_url,
            self._ms_auth_flow,
        )

        if not success:
            return self.async_show_form(
                step_id="ms_credentials",
                data_schema=_STEP_MS_CREDENTIALS,
                errors={"base": "ms_auth_failed"},
            )

        self._ms_token = self._ms_backend.get_token_dict()
        if self._reauth_entry is not None:
            # Reauth: MS re-authorized; reuse the existing Google token.
            return self._create_entry()
        return await self.async_step_google_credentials()

    async def async_step_google_credentials(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 3: enter Google OAuth client credentials."""
        errors: dict[str, str] = {}

        if user_input is not None:
            self._google_client_id = user_input[CONF_GOOGLE_CLIENT_ID].strip()
            self._google_client_secret = user_input[CONF_GOOGLE_CLIENT_SECRET].strip()
            return await self.async_step_google_auth_url()

        return self.async_show_form(
            step_id="google_credentials",
            data_schema=_STEP_GOOGLE_CREDENTIALS,
            errors=errors,
        )

    async def async_step_google_auth_url(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 4: generate Google auth URL and wait for the callback.

        Same external-step pattern as Microsoft: the callback view calls
        async_configure, which must only transition to async_external_step_done.
        Token exchange happens in async_step_google_auth_callback.
        """
        if user_input is not None:
            if "error" in user_input:
                _LOGGER.error(
                    "Google OAuth error from callback: %s", user_input["error"]
                )
                self._google_code = ""
            else:
                self._google_code = user_input.get("code", "")
            return self.async_external_step_done(next_step_id="google_auth_callback")

        if not self.hass.config.external_url:
            return self.async_abort(reason="missing_external_url")

        self.hass.http.register_view(GoogleAuthCallbackView(self.hass))

        try:
            redirect_uri = self._build_google_redirect_uri()
            self._google_flow = await self.hass.async_add_executor_job(
                build_auth_flow,
                self._google_client_id,
                self._google_client_secret,
                redirect_uri,
            )
            auth_url, google_state = await self.hass.async_add_executor_job(
                get_auth_url, self._google_flow
            )
        except Exception:
            _LOGGER.exception("Failed to build Google OAuth URL")
            return self.async_show_form(
                step_id="google_credentials",
                data_schema=_STEP_GOOGLE_CREDENTIALS,
                errors={"base": "google_auth_failed"},
            )

        self.hass.data.setdefault(DOMAIN, {}).setdefault(
            DATA_GOOGLE_OAUTH_FLOWS, {}
        )[google_state] = self.flow_id

        return self.async_external_step(
            step_id="google_auth_url",
            url=auth_url,
        )

    async def async_step_google_auth_callback(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 4b: exchange the Google authorization code for tokens."""
        if not self._google_code or self._google_flow is None:
            return self.async_show_form(
                step_id="google_credentials",
                data_schema=_STEP_GOOGLE_CREDENTIALS,
                errors={"base": "google_auth_failed"},
            )

        try:
            self._google_token = await self.hass.async_add_executor_job(
                exchange_code, self._google_flow, self._google_code
            )
        except Exception:
            _LOGGER.exception("Google OAuth code exchange failed")
            return self.async_show_form(
                step_id="google_credentials",
                data_schema=_STEP_GOOGLE_CREDENTIALS,
                errors={"base": "google_auth_failed"},
            )

        return self._create_entry()

    def _create_entry(self) -> ConfigFlowResult:
        data = {
            CONF_MS_CLIENT_ID: self._ms_client_id,
            CONF_MS_CLIENT_SECRET: self._ms_client_secret,
            "ms_token": self._ms_token,
            CONF_GOOGLE_CLIENT_ID: self._google_client_id,
            CONF_GOOGLE_CLIENT_SECRET: self._google_client_secret,
            "google_token": self._google_token,
        }
        if self._reauth_entry is not None:
            return self.async_update_reload_and_abort(
                self._reauth_entry,
                data=data,
            )
        return self.async_create_entry(
            title="Google → Outlook Contacts",
            data=data,
        )

    def _build_redirect_uri(self) -> str:
        external_url: str = self.hass.config.external_url or ""
        return build_redirect_uri(external_url)

    def _build_google_redirect_uri(self) -> str:
        external_url: str = self.hass.config.external_url or ""
        return external_url.rstrip("/") + GOOGLE_AUTH_CALLBACK_PATH

    @staticmethod
    @callback
    def async_get_options_flow(config_entry: ConfigEntry) -> ContactSyncOptionsFlow:
        return ContactSyncOptionsFlow(config_entry)


class ContactSyncOptionsFlow(OptionsFlow):
    """Allow the user to change sync interval and deletion behaviour."""

    def __init__(self, config_entry: ConfigEntry) -> None:
        self._entry = config_entry

    async def async_step_init(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        if user_input is not None:
            return self.async_create_entry(title="", data=user_input)

        current_options = self._entry.options
        schema = vol.Schema(
            {
                vol.Required(
                    CONF_SYNC_INTERVAL_HOURS,
                    default=current_options.get(
                        CONF_SYNC_INTERVAL_HOURS, DEFAULT_SYNC_INTERVAL_HOURS
                    ),
                ): vol.All(int, vol.Range(min=1, max=168)),
                vol.Required(
                    CONF_DELETE_REMOVED,
                    default=current_options.get(
                        CONF_DELETE_REMOVED, DEFAULT_DELETE_REMOVED
                    ),
                ): bool,
                vol.Required(
                    CONF_AUTO_REMOVE_DUPLICATES,
                    default=current_options.get(
                        CONF_AUTO_REMOVE_DUPLICATES,
                        DEFAULT_AUTO_REMOVE_DUPLICATES,
                    ),
                ): bool,
            }
        )

        return self.async_show_form(step_id="init", data_schema=schema)
