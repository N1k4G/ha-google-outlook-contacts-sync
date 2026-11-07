"""Config flow and options flow for Google Outlook Contacts Sync."""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Callable
from typing import Any

import voluptuous as vol
from homeassistant.config_entries import (
    ConfigEntry,
    ConfigFlow,
    ConfigFlowResult,
    OptionsFlowWithReload,
)
from homeassistant.const import EVENT_HOMEASSISTANT_STOP
from homeassistant.core import Event, callback
from homeassistant.helpers.selector import (
    SelectSelector,
    SelectSelectorConfig,
    SelectSelectorMode,
    TextSelector,
    TextSelectorConfig,
    TextSelectorType,
)

from .auth import CredentialStatus
from .auth.google_auth import build_auth_flow, exchange_code, get_auth_url
from .auth.google_auth import check_credentials as google_check_credentials
from .auth.ms_auth import (
    build_account,
    build_redirect_uri,
    request_token,
)
from .auth.ms_auth import (
    check_credentials as ms_check_credentials,
)
from .auth.ms_auth import (
    get_auth_url as ms_get_auth_url,
)
from .auth.ms_device import DeviceAuthorization, DeviceLoginError, start_device_login
from .const import (
    CONF_AUTO_REMOVE_DUPLICATES,
    CONF_DELETE_REMOVED,
    CONF_GOOGLE_CLIENT_ID,
    CONF_GOOGLE_CLIENT_SECRET,
    CONF_MS_AUTH_MODE,
    CONF_MS_CLIENT_ID,
    CONF_MS_CLIENT_SECRET,
    CONF_SYNC_INTERVAL_HOURS,
    DATA_GOOGLE_OAUTH_FLOWS,
    DATA_MS_OAUTH_FLOWS,
    DATA_REAUTH_PROVIDER,
    DEFAULT_AUTO_REMOVE_DUPLICATES,
    DEFAULT_DELETE_REMOVED,
    DEFAULT_SYNC_INTERVAL_HOURS,
    DOMAIN,
    GOOGLE_AUTH_CALLBACK_PATH,
    MS_AUTH_CONFIDENTIAL,
    MS_AUTH_DEVICE_CODE,
    MS_AUTH_MODES,
)
from .http import async_register_callback_views

_LOGGER = logging.getLogger(__name__)

# Secrets are masked and never prefilled; in reauth and reconfigure an empty
# secret field keeps the stored secret.
_SECRET_SELECTOR = TextSelector(TextSelectorConfig(type=TextSelectorType.PASSWORD))

_MS_MODE_SELECTOR = SelectSelector(
    SelectSelectorConfig(
        options=list(MS_AUTH_MODES),
        translation_key="ms_auth_mode",
        mode=SelectSelectorMode.DROPDOWN,
    )
)

_STEP_MS_CREDENTIALS = vol.Schema(
    {
        vol.Required(CONF_MS_AUTH_MODE, default=MS_AUTH_DEVICE_CODE): _MS_MODE_SELECTOR,
        vol.Required(CONF_MS_CLIENT_ID): str,
        vol.Optional(CONF_MS_CLIENT_SECRET): _SECRET_SELECTOR,
    }
)

_STEP_GOOGLE_CREDENTIALS = vol.Schema(
    {
        vol.Required(CONF_GOOGLE_CLIENT_ID): str,
        vol.Required(CONF_GOOGLE_CLIENT_SECRET): _SECRET_SELECTOR,
    }
)

_STEP_RECONFIGURE = vol.Schema(
    {
        vol.Required(CONF_MS_AUTH_MODE): _MS_MODE_SELECTOR,
        vol.Required(CONF_MS_CLIENT_ID): str,
        vol.Optional(CONF_MS_CLIENT_SECRET): _SECRET_SELECTOR,
        vol.Required(CONF_GOOGLE_CLIENT_ID): str,
        vol.Optional(CONF_GOOGLE_CLIENT_SECRET): _SECRET_SELECTOR,
    }
)

_STEP_REAUTH_MS = vol.Schema({vol.Optional(CONF_MS_CLIENT_SECRET): _SECRET_SELECTOR})

_STEP_REAUTH_GOOGLE = vol.Schema(
    {vol.Optional(CONF_GOOGLE_CLIENT_SECRET): _SECRET_SELECTOR}
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
        self._ms_auth_mode = MS_AUTH_DEVICE_CODE
        self._device: DeviceAuthorization | None = None
        self._device_task: asyncio.Task[dict[str, Any]] | None = None
        self._device_unsub: Callable[[], None] | None = None
        self._device_unload_registered = False
        self._removed = False
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
        # The entry a reauth or reconfigure flow updates instead of creating one.
        self._existing_entry: ConfigEntry | None = None
        # Its data when the flow started, to write back only what changed.
        self._loaded_data: dict[str, Any] = {}
        # Set by reconfigure when a stored token cannot be kept.
        self._ms_needs_oauth: bool = False
        self._google_needs_oauth: bool = False

    async def async_step_user(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 1: enter Microsoft app credentials."""
        if self._async_current_entries():
            return self.async_abort(reason="already_configured")

        return await self.async_step_ms_credentials(user_input)

    async def async_step_reauth(self, entry_data: dict[str, Any]) -> ConfigFlowResult:
        """Re-authorize the provider whose token is invalid or expired."""
        entry_id = self.context.get("entry_id", "")
        reauth_entry = self.hass.config_entries.async_get_entry(entry_id)
        if reauth_entry is None:
            return await self.async_step_user()
        self._load_existing_entry(reauth_entry)
        if self._ms_auth_mode not in MS_AUTH_MODES:
            return self.async_abort(reason="ms_invalid_mode")

        failed_provider = (
            self.hass.data.get(DOMAIN, {})
            .get(DATA_REAUTH_PROVIDER, {})
            .get(entry_id, "google")
        )
        if failed_provider == "ms":
            return await self.async_step_reauth_ms_confirm()
        return await self.async_step_reauth_confirm()

    async def async_step_reauth_ms_confirm(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Offer a new Microsoft secret, then re-authorize Microsoft.

        A revoked or expired secret is a common cause of the failure; reusing
        it would make the code exchange fail the same way again.
        """
        if self._ms_auth_mode == MS_AUTH_DEVICE_CODE:
            return await self.async_step_ms_device()
        if user_input is not None:
            new_secret = _field(user_input, CONF_MS_CLIENT_SECRET)
            if new_secret:
                self._ms_client_secret = new_secret
            return await self.async_step_ms_auth_url()
        return self.async_show_form(
            step_id="reauth_ms_confirm",
            data_schema=_STEP_REAUTH_MS,
        )

    async def async_step_reauth_confirm(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Offer a new Google secret, then re-authorize Google."""
        if user_input is not None:
            new_secret = _field(user_input, CONF_GOOGLE_CLIENT_SECRET)
            if new_secret:
                self._google_client_secret = new_secret
            return await self.async_step_google_auth_url()
        return self.async_show_form(
            step_id="reauth_confirm",
            data_schema=_STEP_REAUTH_GOOGLE,
        )

    async def async_step_reconfigure(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Change client IDs and secrets while keeping tokens and the mapping.

        A changed secret is validated by redeeming the stored refresh token
        with it before anything is saved. Only a changed client ID, or a token
        the provider no longer accepts, needs a fresh sign-in.
        """
        entry = self._get_reconfigure_entry()
        errors: dict[str, str] = {}

        if user_input is not None:
            self._load_existing_entry(entry)
            errors = await self._async_stage_reconfigure(user_input)
            if not errors:
                if self._ms_needs_oauth:
                    return await self.async_step_ms_auth_url()
                if self._google_needs_oauth:
                    return await self.async_step_google_auth_url()
                return self._create_entry()

        shown = user_input or entry.data
        return self.async_show_form(
            step_id="reconfigure",
            data_schema=self.add_suggested_values_to_schema(
                _STEP_RECONFIGURE,
                {
                    CONF_MS_AUTH_MODE: shown.get(
                        CONF_MS_AUTH_MODE, MS_AUTH_CONFIDENTIAL
                    ),
                    CONF_MS_CLIENT_ID: shown.get(CONF_MS_CLIENT_ID, ""),
                    CONF_GOOGLE_CLIENT_ID: shown.get(CONF_GOOGLE_CLIENT_ID, ""),
                },
            ),
            errors=errors,
        )

    async def _async_stage_reconfigure(
        self, user_input: dict[str, Any]
    ) -> dict[str, str]:
        """Validate the submitted credentials and stage them on the flow.

        Returns form errors. Without errors the flow holds the new credentials
        and tokens, and the *_needs_oauth flags name the providers that still
        need a fresh sign-in. Nothing is written to the entry here.
        """
        ms_mode = user_input.get(CONF_MS_AUTH_MODE, self._ms_auth_mode)
        if ms_mode not in MS_AUTH_MODES:
            return {"base": "ms_invalid_mode"}
        mode_changed = ms_mode != self._ms_auth_mode
        ms_id = _field(user_input, CONF_MS_CLIENT_ID)
        ms_secret = _field(user_input, CONF_MS_CLIENT_SECRET)
        google_id = _field(user_input, CONF_GOOGLE_CLIENT_ID)
        google_secret = _field(user_input, CONF_GOOGLE_CLIENT_SECRET)
        if not ms_id or not google_id:
            return {"base": "missing_client_id"}

        ms_id_changed = ms_id != self._ms_client_id
        google_id_changed = google_id != self._google_client_id
        # A secret belongs to its client; the old one cannot work for a new ID.
        if (
            ms_mode == MS_AUTH_CONFIDENTIAL
            and (ms_id_changed or mode_changed)
            and not ms_secret
        ) or (google_id_changed and not google_secret):
            return {"base": "new_client_requires_secret"}
        ms_secret = (
            ""
            if ms_mode == MS_AUTH_DEVICE_CODE
            else ms_secret or self._ms_client_secret
        )
        google_secret = google_secret or self._google_client_secret

        # A refresh token is bound to its client ID, so a new ID always needs
        # a sign-in; a new secret for the same ID is tried against the token.
        self._ms_needs_oauth = ms_id_changed or mode_changed
        if (
            ms_mode == MS_AUTH_CONFIDENTIAL
            and not self._ms_needs_oauth
            and ms_secret != self._ms_client_secret
        ):
            check = await self.hass.async_add_executor_job(
                ms_check_credentials, ms_id, ms_secret, self._ms_token
            )
            if check.status is CredentialStatus.INVALID_CLIENT:
                return {"base": "ms_invalid_client"}
            if check.status is CredentialStatus.CANNOT_CONNECT:
                return {"base": "cannot_connect"}
            if check.status is CredentialStatus.VALID:
                self._ms_token = check.token
            else:
                self._ms_needs_oauth = True

        self._google_needs_oauth = google_id_changed
        if not google_id_changed and google_secret != self._google_client_secret:
            check = await self.hass.async_add_executor_job(
                google_check_credentials, google_id, google_secret, self._google_token
            )
            if check.status is CredentialStatus.INVALID_CLIENT:
                return {"base": "google_invalid_client"}
            if check.status is CredentialStatus.CANNOT_CONNECT:
                return {"base": "cannot_connect"}
            if check.status is CredentialStatus.VALID:
                self._google_token = check.token
            else:
                self._google_needs_oauth = True

        self._ms_auth_mode = ms_mode
        self._ms_client_id = ms_id
        self._ms_client_secret = ms_secret
        self._google_client_id = google_id
        self._google_client_secret = google_secret
        return {}

    def _load_existing_entry(self, entry: ConfigEntry) -> None:
        """Start from the stored credentials and tokens of *entry*."""
        self._existing_entry = entry
        config_data = entry.data
        self._loaded_data = dict(config_data)
        self._ms_auth_mode = config_data.get(CONF_MS_AUTH_MODE, MS_AUTH_CONFIDENTIAL)
        self._ms_client_id = config_data.get(CONF_MS_CLIENT_ID, "")
        self._ms_client_secret = config_data.get(CONF_MS_CLIENT_SECRET, "")
        self._ms_token = dict(config_data.get("ms_token") or {})
        self._google_client_id = config_data.get(CONF_GOOGLE_CLIENT_ID, "")
        self._google_client_secret = config_data.get(CONF_GOOGLE_CLIENT_SECRET, "")
        self._google_token = dict(config_data.get("google_token") or {})
        self._ms_needs_oauth = False
        self._google_needs_oauth = False

    async def async_step_ms_credentials(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        errors: dict[str, str] = {}

        if user_input is not None:
            # Real forms always submit the mode. Missing mode preserves older
            # callers using the existing confidential credentials step.
            self._ms_auth_mode = user_input.get(CONF_MS_AUTH_MODE, MS_AUTH_CONFIDENTIAL)
            if self._ms_auth_mode not in MS_AUTH_MODES:
                return self._ms_credentials_form({"base": "ms_invalid_mode"})
            self._ms_client_id = _field(user_input, CONF_MS_CLIENT_ID)
            self._ms_client_secret = (
                ""
                if self._ms_auth_mode == MS_AUTH_DEVICE_CODE
                else _field(user_input, CONF_MS_CLIENT_SECRET)
            )
            if not self._ms_client_id:
                return self._ms_credentials_form({"base": "missing_client_id"})
            if (
                self._ms_auth_mode == MS_AUTH_CONFIDENTIAL
                and not self._ms_client_secret
            ):
                return self._ms_credentials_form({"base": "new_client_requires_secret"})
            return await self.async_step_ms_auth_url()

        return self._ms_credentials_form(errors)

    async def async_step_ms_auth_url(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 2: generate MS auth URL and wait for the callback.

        HA requires that an external step only transitions to
        async_external_step_done — not to a form. Token exchange therefore
        happens in async_step_ms_auth_callback.
        """
        if self._ms_auth_mode == MS_AUTH_DEVICE_CODE:
            return await self.async_step_ms_device()
        if user_input is not None:
            if "error" in user_input:
                _LOGGER.error("MS OAuth error from callback: %s", user_input["error"])
                self._ms_authorization_url = ""
            else:
                self._ms_authorization_url = user_input.get("authorization_url", "")
            return self.async_external_step_done(next_step_id="ms_auth_callback")

        if not self.hass.config.external_url:
            return self.async_abort(reason="missing_external_url")

        # async_setup has not run during first-time setup (no config entry
        # exists yet), so the callback views are registered here; the helper
        # makes repeated flow attempts idempotent.
        async_register_callback_views(self.hass)

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
            return self._ms_credentials_form({"base": "ms_url_failed"})

        msal_state: str = (
            self._ms_auth_flow.get("state", "") if self._ms_auth_flow else ""
        )
        self.hass.data.setdefault(DOMAIN, {}).setdefault(DATA_MS_OAUTH_FLOWS, {})[
            msal_state
        ] = self.flow_id

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
            return self._ms_credentials_form({"base": "ms_auth_failed"})

        success = await self.hass.async_add_executor_job(
            request_token,
            self._ms_account,
            self._ms_authorization_url,
            self._ms_auth_flow,
        )

        if not success:
            return self._ms_credentials_form({"base": "ms_auth_failed"})

        self._ms_token = self._ms_backend.get_token_dict()
        if self._existing_entry is not None:
            # Reauth or reconfigure: keep the existing Google token unless
            # reconfigure found it unusable.
            if self._google_needs_oauth:
                return await self.async_step_google_auth_url()
            return self._create_entry()
        return await self.async_step_google_credentials()

    async def async_step_ms_device(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Show the code while a cancellable task polls Microsoft in the executor."""
        if self._removed:
            return self.async_abort(reason="ms_device_cancelled")
        if self._device_task is None:
            try:
                async with asyncio.timeout(30):
                    device = await self.hass.async_add_executor_job(
                        start_device_login, self._ms_client_id
                    )
            except DeviceLoginError as exc:
                return self._device_failure(str(exc))
            except Exception:
                return self._device_failure("cannot_connect")
            if self._removed:
                device.cancel()
                return self.async_abort(reason="ms_device_cancelled")
            self._device = device
            self._device_task = self.hass.async_create_background_task(
                device.async_wait(self.hass.async_add_executor_job),
                "Microsoft device authorization",
            )
            self._device_unsub = self.hass.bus.async_listen_once(
                EVENT_HOMEASSISTANT_STOP, self._stop_device
            )
            if self._existing_entry is not None and not self._device_unload_registered:
                self._existing_entry.async_on_unload(self._stop_device)
                self._device_unload_registered = True
        if self._device_task.done():
            return self.async_show_progress_done(next_step_id="ms_device_complete")
        assert self._device is not None
        return self.async_show_progress(
            step_id="ms_device",
            progress_action="ms_device",
            description_placeholders=self._device.placeholders,
            progress_task=self._device_task,
        )

    async def async_step_ms_device_complete(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Stage a verified token only after the progress step has completed."""
        if self._removed or self._device_task is None or self._device_task.cancelled():
            return self.async_abort(reason="ms_device_cancelled")
        if not self._device_task.done():
            return await self.async_step_ms_device()
        try:
            token = self._device_task.result()
        except DeviceLoginError as exc:
            return self._device_failure(str(exc))
        except Exception:
            return self._device_failure("cannot_connect")
        previous_accounts = _ms_account_ids(self._loaded_data.get("ms_token"))
        if previous_accounts and not previous_accounts.intersection(
            _ms_account_ids(token)
        ):
            # Mapping IDs belong to the original Outlook account. A different
            # account must never silently inherit them during migration/reauth.
            return self._device_failure("ms_device_account_mismatch")
        self._ms_token = token
        self._clear_device()
        if self._existing_entry is not None:
            if self._google_needs_oauth:
                return await self.async_step_google_auth_url()
            return self._create_entry()
        return await self.async_step_google_credentials()

    async def async_step_ms_device_failed(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Let the user explicitly start a fresh code after a failed attempt."""
        if user_input is not None:
            return await self.async_step_ms_device()
        return self._device_failure("ms_device_failed")

    def _device_failure(self, reason: str) -> ConfigFlowResult:
        self._clear_device()
        if self._removed:
            return self.async_abort(reason="ms_device_cancelled")
        return self.async_show_form(
            step_id="ms_device_failed",
            data_schema=vol.Schema({}),
            errors={"base": reason},
        )

    @callback
    def _clear_device(self) -> None:
        if self._device is not None:
            self._device.cancel()
        if self._device_task is not None:
            if not self._device_task.done():
                self._device_task.cancel()
            elif not self._device_task.cancelled():
                # Consume failures even if the user closes the flow just as a
                # request completes; no unhandled task exception may leak.
                self._device_task.exception()
        self._device = None
        self._device_task = None
        if self._device_unsub is not None:
            self._device_unsub()
            self._device_unsub = None

    @callback
    def _stop_device(self, event: Event | None = None) -> None:
        if event is not None:
            # HA already removed this one-shot listener before invoking us.
            self._device_unsub = None
        self._removed = True
        self._clear_device()

    @callback
    def async_remove(self) -> None:
        """Discard ephemeral credentials and prevent late results being saved."""
        self._stop_device()
        super().async_remove()

    async def async_step_google_credentials(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 3: enter Google OAuth client credentials."""
        errors: dict[str, str] = {}

        if user_input is not None:
            self._google_client_id = user_input[CONF_GOOGLE_CLIENT_ID].strip()
            self._google_client_secret = user_input[CONF_GOOGLE_CLIENT_SECRET].strip()
            return await self.async_step_google_auth_url()

        return self._google_credentials_form(errors)

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

        async_register_callback_views(self.hass)

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
            return self._google_credentials_form({"base": "google_auth_failed"})

        self.hass.data.setdefault(DOMAIN, {}).setdefault(DATA_GOOGLE_OAUTH_FLOWS, {})[
            google_state
        ] = self.flow_id

        return self.async_external_step(
            step_id="google_auth_url",
            url=auth_url,
        )

    async def async_step_google_auth_callback(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Step 4b: exchange the Google authorization code for tokens."""
        if not self._google_code or self._google_flow is None:
            return self._google_credentials_form({"base": "google_auth_failed"})

        try:
            self._google_token = await self.hass.async_add_executor_job(
                exchange_code, self._google_flow, self._google_code
            )
        except Exception:
            _LOGGER.exception("Google OAuth code exchange failed")
            return self._google_credentials_form({"base": "google_auth_failed"})

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
        if self._ms_auth_mode == MS_AUTH_DEVICE_CODE:
            data.pop(CONF_MS_CLIENT_SECRET)
        if (
            self._existing_entry is None
            or CONF_MS_AUTH_MODE in self._loaded_data
            or self._ms_auth_mode != MS_AUTH_CONFIDENTIAL
        ):
            data[CONF_MS_AUTH_MODE] = self._ms_auth_mode
        if self._existing_entry is not None:
            # Only what this flow changed. A sign-in can take minutes, during
            # which the running integration may persist a refreshed token;
            # writing back the copy loaded at the start would undo that. Keys
            # the flow does not own (a legacy delete_removed, say) stay too.
            changed = {
                key: value
                for key, value in data.items()
                if value != self._loaded_data.get(key)
            }
            # Merge against the current entry so a concurrent Google refresh
            # is retained, then explicitly remove the obsolete Microsoft secret.
            merged = {**self._existing_entry.data, **changed}
            if self._ms_auth_mode == MS_AUTH_DEVICE_CODE:
                merged.pop(CONF_MS_CLIENT_SECRET, None)
            return self.async_update_reload_and_abort(
                self._existing_entry,
                data=merged,
            )
        return self.async_create_entry(
            title="Google → Outlook Contacts",
            data=data,
        )

    def _ms_credentials_form(self, errors: dict[str, str]) -> ConfigFlowResult:
        """The Microsoft credentials form, also the retry form after a failure.

        The client ID is kept from the attempt (or the entry being reauthorized
        or reconfigured), so a retry only needs the secret again.
        """
        return self.async_show_form(
            step_id="ms_credentials",
            data_schema=self.add_suggested_values_to_schema(
                _STEP_MS_CREDENTIALS,
                {
                    CONF_MS_CLIENT_ID: self._ms_client_id,
                    CONF_MS_AUTH_MODE: self._ms_auth_mode,
                },
            ),
            errors=errors,
        )

    def _google_credentials_form(self, errors: dict[str, str]) -> ConfigFlowResult:
        """The Google credentials form, also the retry form after a failure."""
        return self.async_show_form(
            step_id="google_credentials",
            data_schema=self.add_suggested_values_to_schema(
                _STEP_GOOGLE_CREDENTIALS,
                {CONF_GOOGLE_CLIENT_ID: self._google_client_id},
            ),
            errors=errors,
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


def _ms_account_ids(token: dict[str, Any] | None) -> set[str]:
    """Read account identity, never token contents, for mapping-safe migration."""
    accounts = (token or {}).get("Account", {})
    if not isinstance(accounts, dict):
        return set()
    return {
        account["home_account_id"]
        for account in accounts.values()
        if isinstance(account, dict)
        and isinstance(account.get("home_account_id"), str)
        and account["home_account_id"]
    }


def _field(user_input: dict[str, Any], key: str) -> str:
    """A submitted text field, stripped; an omitted optional field is empty."""
    return str(user_input.get(key) or "").strip()


class ContactSyncOptionsFlow(OptionsFlowWithReload):
    """Allow the user to change sync interval and deletion behaviour.

    OptionsFlowWithReload reloads the entry only when the options actually
    changed, which replaces the update listener this integration used to have.
    """

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
