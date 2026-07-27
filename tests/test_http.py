"""Tests for http.py — the unauthenticated OAuth callback views."""

from __future__ import annotations

from http import HTTPStatus
from unittest.mock import AsyncMock, MagicMock

import pytest

from custom_components.google_outlook_contacts_sync.const import (
    DATA_GOOGLE_OAUTH_FLOWS,
    DATA_MS_OAUTH_FLOWS,
    DOMAIN,
)
from custom_components.google_outlook_contacts_sync.http import (
    GoogleAuthCallbackView,
    MSAuthCallbackView,
)


def _hass(flow_key: str, states: dict[str, str] | None = None) -> MagicMock:
    hass = MagicMock()
    hass.data = {DOMAIN: {flow_key: dict(states or {})}}
    hass.config_entries.flow.async_configure = AsyncMock()
    return hass


def _request(query: dict[str, str], url: str = "https://ha.example/cb") -> MagicMock:
    request = MagicMock()
    request.query = query
    request.url = url
    return request


class TestMicrosoftCallback:
    @pytest.mark.asyncio
    async def test_valid_state_resumes_the_flow(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        response = await view.get(
            _request({"state": "state-1", "code": "abc"}, "https://ha.example/cb?code=abc")
        )

        assert response.status == HTTPStatus.OK
        kwargs = hass.config_entries.flow.async_configure.call_args.kwargs
        assert kwargs["flow_id"] == "flow-1"
        assert kwargs["user_input"]["authorization_url"] == "https://ha.example/cb?code=abc"

    @pytest.mark.asyncio
    async def test_unknown_state_is_rejected(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        response = await view.get(_request({"state": "nope"}))

        assert response.status == HTTPStatus.BAD_REQUEST
        hass.config_entries.flow.async_configure.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_state_is_rejected(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        assert (await view.get(_request({}))).status == HTTPStatus.BAD_REQUEST

    @pytest.mark.asyncio
    async def test_state_is_consumed_so_a_replay_fails(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)
        request = _request({"state": "state-1", "code": "abc"})

        first = await view.get(request)
        replay = await view.get(request)

        assert first.status == HTTPStatus.OK
        assert replay.status == HTTPStatus.BAD_REQUEST

    @pytest.mark.asyncio
    async def test_oauth_error_is_forwarded_to_the_flow(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        response = await view.get(
            _request({"state": "state-1", "error": "access_denied"})
        )

        assert response.status == HTTPStatus.OK
        kwargs = hass.config_entries.flow.async_configure.call_args.kwargs
        assert kwargs["user_input"]["error"] == "access_denied"


class TestGoogleCallback:
    @pytest.mark.asyncio
    async def test_valid_state_forwards_the_code(self):
        hass = _hass(DATA_GOOGLE_OAUTH_FLOWS, {"state-g": "flow-g"})
        view = GoogleAuthCallbackView(hass)

        response = await view.get(_request({"state": "state-g", "code": "xyz"}))

        assert response.status == HTTPStatus.OK
        kwargs = hass.config_entries.flow.async_configure.call_args.kwargs
        assert kwargs["user_input"]["code"] == "xyz"

    @pytest.mark.asyncio
    async def test_unknown_state_is_rejected(self):
        hass = _hass(DATA_GOOGLE_OAUTH_FLOWS, {})
        view = GoogleAuthCallbackView(hass)

        assert (
            await view.get(_request({"state": "nope"}))
        ).status == HTTPStatus.BAD_REQUEST

    @pytest.mark.asyncio
    async def test_state_is_consumed_so_a_replay_fails(self):
        hass = _hass(DATA_GOOGLE_OAUTH_FLOWS, {"state-g": "flow-g"})
        view = GoogleAuthCallbackView(hass)
        request = _request({"state": "state-g", "code": "xyz"})

        assert (await view.get(request)).status == HTTPStatus.OK
        assert (await view.get(request)).status == HTTPStatus.BAD_REQUEST


class TestErrorPageEscaping:
    """#13: both views are unauthenticated and echo query parameters back."""

    XSS = "<script>alert(1)</script>"

    @pytest.mark.asyncio
    async def test_microsoft_error_description_is_escaped(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        response = await view.get(
            _request(
                {
                    "state": "state-1",
                    "error": "invalid_request",
                    "error_description": self.XSS,
                }
            )
        )

        assert self.XSS not in response.text
        assert "&lt;script&gt;" in response.text

    @pytest.mark.asyncio
    async def test_google_error_is_escaped(self):
        hass = _hass(DATA_GOOGLE_OAUTH_FLOWS, {"state-g": "flow-g"})
        view = GoogleAuthCallbackView(hass)

        response = await view.get(_request({"state": "state-g", "error": self.XSS}))

        assert self.XSS not in response.text
        assert "&lt;script&gt;" in response.text

    @pytest.mark.asyncio
    async def test_long_error_descriptions_are_truncated(self):
        hass = _hass(DATA_MS_OAUTH_FLOWS, {"state-1": "flow-1"})
        view = MSAuthCallbackView(hass)

        response = await view.get(
            _request(
                {
                    "state": "state-1",
                    "error": "invalid_request",
                    "error_description": "A" * 5000,
                }
            )
        )

        assert len(response.text) < 2000
