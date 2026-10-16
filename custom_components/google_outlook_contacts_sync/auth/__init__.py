"""Auth sub-package."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import Any


class CredentialStatus(StrEnum):
    """Outcome of refreshing a stored token with a given client credential."""

    # The refresh succeeded; the credential works with the stored token.
    VALID = "valid"
    # The provider rejected the client itself (wrong or expired secret).
    INVALID_CLIENT = "invalid_client"
    # The credential may be fine, but the stored token cannot be redeemed with
    # it; only a fresh authorization can help.
    REAUTHORIZE = "reauthorize"
    # The provider could not be reached; nothing can be concluded.
    CANNOT_CONNECT = "cannot_connect"


@dataclass(frozen=True)
class CredentialCheck:
    """Result of a credential check, with the refreshed token on success."""

    status: CredentialStatus
    token: dict[str, Any] | None = None
