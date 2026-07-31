"""Duplicate detection for Outlook contacts.

Duplicates are matched by normalized name only (per configuration). This is
deliberately aggressive: two genuinely different people who share a normalized
name are treated as duplicates. Removal is therefore opt-in and always keeps
exactly one contact per group.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# Fields counted to gauge how "complete" a contact is when choosing a keeper.
_COMPLETENESS_FIELDS = (
    "givenName",
    "surname",
    "displayName",
    "emailAddresses",
    "homePhones",
    "businessPhones",
    "mobilePhone",
    "homeAddress",
    "businessAddress",
    "otherAddress",
    "companyName",
    "jobTitle",
    "birthday",
)

# Properties needed for dedup bookkeeping but not scored for completeness.
_METADATA_FIELDS = ("id", "lastModifiedDateTime")

# Top-level Graph contact properties requested for dedup decisions.
#
# Derived from the counted fields rather than written out separately: when these
# drifted apart, 8 of 12 scored properties were never requested, so the keeper
# heuristic silently degraded to "has a name/email, modified recently" while
# deleting the losing contacts.
SELECT_FIELDS = ",".join(_METADATA_FIELDS + _COMPLETENESS_FIELDS)


def normalize_email(address: str) -> str:
    """Normalize an email address for identity comparison."""
    return address.strip().casefold()


def emails_from_graph(raw: dict[str, Any]) -> frozenset[str]:
    """Normalized email addresses of a raw Graph contact."""
    entries = raw.get("emailAddresses") or []
    if not isinstance(entries, list):
        return frozenset()
    return frozenset(
        normalized
        for entry in entries
        if isinstance(entry, dict)
        and (normalized := normalize_email(str(entry.get("address") or "")))
    )


@dataclass
class OutlookContact:
    """Minimal view of an Outlook contact for dedup and adoption decisions."""

    contact_id: str
    display_name: str = ""
    given_name: str = ""
    surname: str = ""
    field_count: int = 0
    last_modified: str = ""
    emails: frozenset[str] = frozenset()

    @classmethod
    def from_graph(cls, raw: dict[str, Any]) -> OutlookContact:
        return cls(
            contact_id=str(raw.get("id", "")),
            display_name=raw.get("displayName") or "",
            given_name=raw.get("givenName") or "",
            surname=raw.get("surname") or "",
            field_count=sum(1 for f in _COMPLETENESS_FIELDS if raw.get(f)),
            last_modified=raw.get("lastModifiedDateTime") or "",
            emails=emails_from_graph(raw),
        )

    def name_key(self) -> str | None:
        """Return the normalized name used to group duplicates, or None."""
        name = self.display_name or f"{self.given_name} {self.surname}"
        key = " ".join(name.split()).casefold()
        return key or None


@dataclass
class DuplicateGroup:
    """A set of Outlook contacts sharing the same name key."""

    key: str
    contacts: list[OutlookContact] = field(default_factory=list)


def find_duplicate_groups(
    contacts: list[OutlookContact],
) -> list[DuplicateGroup]:
    """Group contacts by name key and return only groups with more than one."""
    groups: dict[str, list[OutlookContact]] = {}
    for contact in contacts:
        key = contact.name_key()
        if key is None:
            # Contacts with no usable name are never treated as duplicates.
            continue
        groups.setdefault(key, []).append(contact)
    return [
        DuplicateGroup(key=key, contacts=members)
        for key, members in groups.items()
        if len(members) > 1
    ]


def choose_keeper(group: DuplicateGroup, mapped_ids: set[str]) -> OutlookContact:
    """Pick the contact to keep from a duplicate group.

    Preference order: a contact already tracked in our mapping, then the most
    complete contact, then the most recently modified.
    """
    return max(
        group.contacts,
        key=lambda c: (
            c.contact_id in mapped_ids,
            c.field_count,
            c.last_modified,
        ),
    )
