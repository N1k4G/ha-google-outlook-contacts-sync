"""Shared pytest fixtures."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def load_fixture(name: str) -> dict[str, Any]:
    """Load a JSON fixture file from tests/fixtures/."""
    return json.loads((FIXTURES_DIR / name).read_text())


@pytest.fixture
def person_full() -> dict[str, Any]:
    return load_fixture("google_person_full.json")


@pytest.fixture
def person_no_year_birthday() -> dict[str, Any]:
    return load_fixture("google_person_no_year_birthday.json")


@pytest.fixture
def person_deleted() -> dict[str, Any]:
    return load_fixture("google_person_deleted.json")
