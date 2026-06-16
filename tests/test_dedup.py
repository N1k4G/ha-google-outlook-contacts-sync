"""Tests for sync/dedup.py — pure functions, no HA required."""

from __future__ import annotations

from custom_components.google_outlook_contacts_sync.sync.dedup import (
    OutlookContact,
    choose_keeper,
    find_duplicate_groups,
)


def _c(
    contact_id: str,
    display_name: str = "",
    given: str = "",
    surname: str = "",
    fields: int = 0,
    modified: str = "",
) -> OutlookContact:
    return OutlookContact(
        contact_id=contact_id,
        display_name=display_name,
        given_name=given,
        surname=surname,
        field_count=fields,
        last_modified=modified,
    )


class TestNameKey:
    def test_uses_display_name_normalized(self) -> None:
        assert _c("1", display_name="  Alice   Smith ").name_key() == "alice smith"

    def test_falls_back_to_given_surname(self) -> None:
        assert _c("1", given="Bob", surname="Jones").name_key() == "bob jones"

    def test_blank_name_is_none(self) -> None:
        assert _c("1").name_key() is None


class TestFindDuplicateGroups:
    def test_groups_only_collisions(self) -> None:
        contacts = [
            _c("1", display_name="Alice Smith"),
            _c("2", display_name="alice  smith"),  # same key, different case/space
            _c("3", display_name="Carol Jones"),
        ]
        groups = find_duplicate_groups(contacts)
        assert len(groups) == 1
        assert groups[0].key == "alice smith"
        assert {c.contact_id for c in groups[0].contacts} == {"1", "2"}

    def test_nameless_contacts_never_grouped(self) -> None:
        contacts = [_c("1"), _c("2")]
        assert find_duplicate_groups(contacts) == []


class TestChooseKeeper:
    def test_prefers_mapped_contact(self) -> None:
        group = find_duplicate_groups(
            [
                _c("1", display_name="Al", fields=1),
                _c("2", display_name="Al", fields=9),  # more complete but unmapped
            ]
        )[0]
        keeper = choose_keeper(group, mapped_ids={"1"})
        assert keeper.contact_id == "1"

    def test_prefers_most_complete_then_recent(self) -> None:
        group = find_duplicate_groups(
            [
                _c("1", display_name="Al", fields=3, modified="2024-01-01"),
                _c("2", display_name="Al", fields=5, modified="2023-01-01"),
            ]
        )[0]
        keeper = choose_keeper(group, mapped_ids=set())
        assert keeper.contact_id == "2"
