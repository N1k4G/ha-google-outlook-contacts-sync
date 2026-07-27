"""HA HTTP views that handle OAuth redirect callbacks."""

from __future__ import annotations

import html
import logging
from http import HTTPStatus

from aiohttp import web
from homeassistant.components.http import (  # type: ignore[attr-defined]
    HomeAssistantView,
)
from homeassistant.core import HomeAssistant

from .const import (
    DATA_GOOGLE_OAUTH_FLOWS,
    DATA_MS_OAUTH_FLOWS,
    DATA_VIEWS_REGISTERED,
    DOMAIN,
    GOOGLE_AUTH_CALLBACK_NAME,
    GOOGLE_AUTH_CALLBACK_PATH,
    MS_AUTH_CALLBACK_NAME,
    MS_AUTH_CALLBACK_PATH,
)

_LOGGER = logging.getLogger(__name__)


def async_register_callback_views(hass: HomeAssistant) -> None:
    """Register both OAuth callback views, at most once per HA run.

    Called from async_setup and from the config flow steps, because during
    first-time setup no config entry exists yet and async_setup has not run.
    Every register_view call adds another route and handler, so repeated flow
    attempts used to accumulate duplicates.
    """
    domain_data = hass.data.setdefault(DOMAIN, {})
    if domain_data.get(DATA_VIEWS_REGISTERED):
        return

    hass.http.register_view(MSAuthCallbackView(hass))
    hass.http.register_view(GoogleAuthCallbackView(hass))
    domain_data[DATA_VIEWS_REGISTERED] = True


class MSAuthCallbackView(HomeAssistantView):
    """Handle the Microsoft OAuth callback redirect."""

    url = MS_AUTH_CALLBACK_PATH
    name = MS_AUTH_CALLBACK_NAME
    requires_auth = False

    def __init__(self, hass: HomeAssistant) -> None:
        self.hass = hass

    async def get(self, request: web.Request) -> web.Response:
        """Handle GET /api/google_outlook_contacts_sync/auth."""
        params = dict(request.query)
        _LOGGER.debug("MS OAuth callback received: %s", list(params.keys()))

        state = params.get("state", "")
        flow_id = (
            self.hass.data.get(DOMAIN, {}).get(DATA_MS_OAUTH_FLOWS, {}).pop(state, None)
        )

        if flow_id is None:
            _LOGGER.error(
                "MS OAuth callback received with unrecognized state: %s", state
            )
            return web.Response(
                status=HTTPStatus.BAD_REQUEST,
                text="Invalid state parameter.",
            )

        if "error" in params:
            _LOGGER.error(
                "MS OAuth error: %s — %s",
                params.get("error"),
                params.get("error_description", ""),
            )
            await self.hass.config_entries.flow.async_configure(
                flow_id=flow_id,
                user_input={"error": params["error"]},
            )
            return web.Response(
                status=HTTPStatus.OK,
                content_type="text/html",
                text=_error_page(params.get("error_description", params["error"])),
            )

        authorization_url = str(request.url)
        await self.hass.config_entries.flow.async_configure(
            flow_id=flow_id,
            user_input={"authorization_url": authorization_url},
        )

        return web.Response(
            status=HTTPStatus.OK,
            content_type="text/html",
            text=_success_page("Microsoft"),
        )


class GoogleAuthCallbackView(HomeAssistantView):
    """Handle the Google OAuth callback redirect."""

    url = GOOGLE_AUTH_CALLBACK_PATH
    name = GOOGLE_AUTH_CALLBACK_NAME
    requires_auth = False

    def __init__(self, hass: HomeAssistant) -> None:
        self.hass = hass

    async def get(self, request: web.Request) -> web.Response:
        """Handle GET /api/google_outlook_contacts_sync/google_auth."""
        params = dict(request.query)
        _LOGGER.debug("Google OAuth callback received: %s", list(params.keys()))

        state = params.get("state", "")
        flow_id = (
            self.hass.data.get(DOMAIN, {})
            .get(DATA_GOOGLE_OAUTH_FLOWS, {})
            .pop(state, None)
        )

        if flow_id is None:
            _LOGGER.error(
                "Google OAuth callback received with unrecognized state: %s", state
            )
            return web.Response(
                status=HTTPStatus.BAD_REQUEST,
                text="Invalid state parameter.",
            )

        if "error" in params:
            _LOGGER.error("Google OAuth error: %s", params.get("error"))
            await self.hass.config_entries.flow.async_configure(
                flow_id=flow_id,
                user_input={"error": params["error"]},
            )
            return web.Response(
                status=HTTPStatus.OK,
                content_type="text/html",
                text=_error_page(params.get("error", "Authorization failed")),
            )

        code = params.get("code", "")
        await self.hass.config_entries.flow.async_configure(
            flow_id=flow_id,
            user_input={"code": code},
        )

        return web.Response(
            status=HTTPStatus.OK,
            content_type="text/html",
            text=_success_page("Google"),
        )


# OAuth error descriptions are attacker-controllable query parameters echoed
# back by an unauthenticated view, so everything interpolated into these pages
# is escaped.
_MAX_ERROR_LENGTH = 500


def _success_page(product: str) -> str:
    return (
        f"<!DOCTYPE html><html><head><title>Authorized</title></head><body>"
        f"<h2>{html.escape(product)} account authorized!</h2>"
        f"<p>You can close this tab and return to Home Assistant.</p>"
        f"</body></html>"
    )


def _error_page(message: str) -> str:
    truncated = message[:_MAX_ERROR_LENGTH]
    return (
        f"<!DOCTYPE html><html><head><title>Error</title></head><body>"
        f"<h2>Authorization failed</h2><p>{html.escape(truncated)}</p>"
        f"<p>Please close this tab and try again in Home Assistant.</p>"
        f"</body></html>"
    )
