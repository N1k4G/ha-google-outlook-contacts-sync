"""Translation files must stay in step with strings.json and with the flow.

A key used by config_flow.py but missing from a translation renders as the raw
key to that user, which is how error.ms_url_failed and abort.missing_external_url
reached English users unnoticed.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

COMPONENT = (
    Path(__file__).parent.parent / "custom_components" / "google_outlook_contacts_sync"
)
STRINGS = COMPONENT / "strings.json"
TRANSLATIONS = sorted((COMPONENT / "translations").glob("*.json"))


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _keys(data: dict, prefix: str = "") -> set[str]:
    found: set[str] = set()
    for key, value in data.items():
        found.add(prefix + key)
        if isinstance(value, dict):
            found |= _keys(value, f"{prefix}{key}.")
    return found


@pytest.mark.parametrize("path", TRANSLATIONS, ids=lambda p: p.name)
def test_translation_matches_strings(path: Path) -> None:
    expected = _keys(_load(STRINGS))
    actual = _keys(_load(path))

    assert not expected - actual, f"{path.name} is missing keys"
    assert not actual - expected, f"{path.name} has keys absent from strings.json"


@pytest.mark.parametrize("section", ["error", "abort"])
def test_every_key_used_by_the_flow_exists(section: str) -> None:
    """Scan config_flow.py for the error/abort keys it actually raises."""
    source = (COMPONENT / "config_flow.py").read_text(encoding="utf-8")

    if section == "error":
        used = set(re.findall(r'"base":\s*"([a-z_]+)"', source))
    else:
        used = set(re.findall(r'async_abort\(\s*reason="([a-z_]+)"', source))

    assert used, f"no {section} keys found in config_flow.py — check the regex"

    for path in [STRINGS, *TRANSLATIONS]:
        available = set(_load(path)["config"].get(section, {}))
        missing = used - available
        assert not missing, f"{path.name} lacks {section} keys: {sorted(missing)}"


def test_google_step_does_not_promise_a_desktop_client() -> None:
    """The flow builds a web OAuth client with a redirect URI (#15)."""
    for path in [STRINGS, *TRANSLATIONS]:
        description = _load(path)["config"]["step"]["google_credentials"]["description"]
        assert "Desktop" not in description
        assert "Desktop-Client" not in description
