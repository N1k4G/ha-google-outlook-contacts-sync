"""Diagnostics must never expose a credential, including a rotated one."""

from __future__ import annotations

import json
from unittest.mock import MagicMock

from custom_components.google_outlook_contacts_sync.const import (
    DATA_COORDINATOR,
    DATA_STORE,
    DOMAIN,
)
from custom_components.google_outlook_contacts_sync.diagnostics import (
    async_get_config_entry_diagnostics,
)


async def test_credentials_and_tokens_are_redacted() -> None:
    secrets = ["mid", "msecret", "ms-rt", "gid", "gsecret", "grefresh", "gaccess"]
    entry = MagicMock()
    entry.entry_id = "entry-1"
    entry.options = {}
    entry.data = {
        "ms_client_id": "mid",
        "ms_client_secret": "msecret",
        "ms_token": {"RefreshToken": {"k": {"secret": "ms-rt"}}},
        "google_client_id": "gid",
        "google_client_secret": "gsecret",
        "google_token": {
            "client_id": "gid",
            "client_secret": "gsecret",
            "refresh_token": "grefresh",
            "token": "gaccess",
        },
    }
    coordinator = MagicMock()
    coordinator.data = None
    coordinator.last_exception = None
    coordinator.update_interval = None
    store = MagicMock()
    store.all_resource_names.return_value = []
    store.sync_token = None
    hass = MagicMock()
    hass.data = {
        DOMAIN: {"entry-1": {DATA_COORDINATOR: coordinator, DATA_STORE: store}}
    }

    result = await async_get_config_entry_diagnostics(hass, entry)

    dumped = json.dumps(result, default=str)
    for secret in secrets:
        assert f'"{secret}"' not in dumped
