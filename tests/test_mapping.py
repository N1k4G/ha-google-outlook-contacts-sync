"""Tests for sync/mapping.py — pure logic, no network."""

from __future__ import annotations

from typing import Any

from custom_components.google_outlook_contacts_sync.const import (
    BIRTHDAY_PLACEHOLDER_YEAR,
)
from custom_components.google_outlook_contacts_sync.sync.mapping import (
    ADDRESS_SUBFIELDS,
    MANAGED_GRAPH_FIELDS,
    contact_hash,
    is_deleted,
    to_graph_contact,
    to_graph_contact_update,
)


class TestToGraphContact:
    def test_full_person_names(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        assert result["givenName"] == "Jane"
        assert result["surname"] == "Example"
        assert result["displayName"] == "Jane Example"
        assert result["fileAs"] == "Jane Example"

    def test_full_person_emails(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        emails = result["emailAddresses"]
        assert len(emails) == 2
        addresses = [e["address"] for e in emails]
        assert "jane.example@personal.example" in addresses
        assert "jane.example@work.example" in addresses

    def test_email_name_is_the_contact_name_not_a_type_label(
        self, person_full: dict[str, Any]
    ) -> None:
        """Graph's emailAddresses[].name is a display name (#29)."""
        result = to_graph_contact(person_full)
        emails = {e["address"]: e["name"] for e in result["emailAddresses"]}
        assert emails["jane.example@personal.example"] == "Jane Example"
        assert emails["jane.example@work.example"] == "Jane Example"

    def test_full_person_phones(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        assert result["mobilePhone"] == "+49 170 1234567"
        assert result["homePhones"] == ["+49 30 12345678"]
        assert "businessPhones" not in result

    def test_full_person_birthday_with_year(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        assert result["birthday"] == "1985-03-15T00:00:00Z"

    def test_birthday_no_year_uses_placeholder(
        self, person_no_year_birthday: dict[str, Any]
    ) -> None:
        result = to_graph_contact(person_no_year_birthday)
        assert result["birthday"] == f"{BIRTHDAY_PLACEHOLDER_YEAR}-07-04T00:00:00Z"

    def test_full_person_address(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        addr = result["homeAddress"]
        assert addr["street"] == "Musterstraße 1"
        assert addr["city"] == "Berlin"
        assert addr["postalCode"] == "10115"
        assert addr["countryOrRegion"] == "DE"

    def test_full_person_organization(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        assert result["companyName"] == "Example Corp"
        assert result["jobTitle"] == "Software Engineer"
        assert result["department"] == "Engineering"

    def test_empty_person_returns_empty_dict(self) -> None:
        result = to_graph_contact({})
        assert result == {}

    def test_person_with_only_name(self) -> None:
        person = {
            "names": [
                {
                    "givenName": "Solo",
                    "familyName": "Name",
                    "displayName": "Solo Name",
                }
            ]
        }
        result = to_graph_contact(person)
        assert result["givenName"] == "Solo"
        assert "emailAddresses" not in result
        assert "mobilePhone" not in result

    def test_skips_empty_email_addresses(self) -> None:
        person = {
            "emailAddresses": [
                {"value": "", "type": "home"},
                {"value": "valid@example.com", "type": "work"},
            ]
        }
        result = to_graph_contact(person)
        assert len(result["emailAddresses"]) == 1
        assert result["emailAddresses"][0]["address"] == "valid@example.com"

    def test_skips_empty_phone_numbers(self) -> None:
        person = {
            "phoneNumbers": [
                {"value": "", "type": "mobile"},
                {"value": "+49123", "type": "home"},
            ]
        }
        result = to_graph_contact(person)
        assert "mobilePhone" not in result
        assert result["homePhones"] == ["+49123"]

    def test_birthday_missing_month_skipped(self) -> None:
        person = {"birthdays": [{"metadata": {"primary": True}, "date": {"day": 5}}]}
        result = to_graph_contact(person)
        assert "birthday" not in result

    def test_birthday_invalid_date_skipped(self) -> None:
        person = {
            "birthdays": [
                {
                    "metadata": {"primary": True},
                    "date": {"year": 2000, "month": 2, "day": 30},
                }
            ]
        }
        result = to_graph_contact(person)
        assert "birthday" not in result

    def test_primary_name_preferred(self) -> None:
        person = {
            "names": [
                {"givenName": "Secondary"},
                {
                    "metadata": {"primary": True},
                    "givenName": "Primary",
                    "displayName": "Primary",
                },
            ]
        }
        result = to_graph_contact(person)
        assert result["givenName"] == "Primary"

    def test_business_phone_type(self) -> None:
        person = {
            "phoneNumbers": [
                {"value": "+49 30 99887766", "type": "work"},
                {"value": "+49 30 11223344", "type": "company_main"},
            ]
        }
        result = to_graph_contact(person)
        assert result["businessPhones"] == ["+49 30 99887766", "+49 30 11223344"]
        assert "homePhones" not in result

    def test_multiple_mobile_only_first_kept(self) -> None:
        person = {
            "phoneNumbers": [
                {"value": "+49 170 1111111", "type": "mobile"},
                {"value": "+49 170 2222222", "type": "mobile"},
            ]
        }
        result = to_graph_contact(person)
        assert result["mobilePhone"] == "+49 170 1111111"

    def test_email_without_contact_name_omits_name(self) -> None:
        person = {"emailAddresses": [{"value": "x@example.com", "type": "fax"}]}
        result = to_graph_contact(person)
        assert result["emailAddresses"][0] == {"address": "x@example.com"}

    def test_google_email_types_never_become_name_labels(self) -> None:
        person = {
            "names": [{"displayName": "Real Person"}],
            "emailAddresses": [
                {"value": "h@example.com", "type": "home"},
                {"value": "w@example.com", "type": "work"},
                {"value": "o@example.com", "type": "other"},
            ],
        }
        result = to_graph_contact(person)
        names = {e.get("name") for e in result["emailAddresses"]}
        assert names == {"Real Person"}
        assert not names & {"personal", "work", "other"}

    def test_email_name_falls_back_to_given_and_family(self) -> None:
        person = {
            "names": [{"givenName": "No", "familyName": "Display"}],
            "emailAddresses": [{"value": "x@example.com"}],
        }
        result = to_graph_contact(person)
        assert result["emailAddresses"][0]["name"] == "No Display"

    def test_address_business_type(self) -> None:
        person = {
            "addresses": [
                {
                    "type": "work",
                    "streetAddress": "Work St 1",
                    "city": "Munich",
                    "postalCode": "80333",
                    "countryCode": "DE",
                }
            ]
        }
        result = to_graph_contact(person)
        assert "businessAddress" in result
        assert result["businessAddress"]["city"] == "Munich"

    def test_only_organization_name_when_no_title(self) -> None:
        person = {"organizations": [{"name": "ACME"}]}
        result = to_graph_contact(person)
        assert result["companyName"] == "ACME"
        assert "jobTitle" not in result


def _addr(addr_type: str, city: str) -> dict[str, Any]:
    return {"type": addr_type, "streetAddress": f"{city} St 1", "city": city}


class TestAddressTypes:
    def test_home_address(self) -> None:
        result = to_graph_contact({"addresses": [_addr("home", "Berlin")]})
        assert result["homeAddress"]["city"] == "Berlin"
        assert "businessAddress" not in result
        assert "otherAddress" not in result

    def test_work_address(self) -> None:
        result = to_graph_contact({"addresses": [_addr("work", "Munich")]})
        assert result["businessAddress"]["city"] == "Munich"
        assert "homeAddress" not in result
        assert "otherAddress" not in result

    def test_other_address_is_not_business(self) -> None:
        """Regression for #32: 'other' used to land in businessAddress."""
        result = to_graph_contact({"addresses": [_addr("other", "Hamburg")]})
        assert result["otherAddress"]["city"] == "Hamburg"
        assert "businessAddress" not in result

    def test_unknown_type_goes_to_other(self) -> None:
        result = to_graph_contact({"addresses": [_addr("vacation", "Sylt")]})
        assert result["otherAddress"]["city"] == "Sylt"
        assert "businessAddress" not in result

    def test_all_three_slots_populate_independently(self) -> None:
        result = to_graph_contact(
            {
                "addresses": [
                    _addr("home", "Berlin"),
                    _addr("work", "Munich"),
                    _addr("other", "Hamburg"),
                ]
            }
        )
        assert result["homeAddress"]["city"] == "Berlin"
        assert result["businessAddress"]["city"] == "Munich"
        assert result["otherAddress"]["city"] == "Hamburg"

    def test_first_address_per_slot_wins(self) -> None:
        result = to_graph_contact(
            {"addresses": [_addr("home", "First"), _addr("home", "Second")]}
        )
        assert result["homeAddress"]["city"] == "First"

    def test_empty_address_does_not_consume_the_slot(self) -> None:
        result = to_graph_contact(
            {"addresses": [{"type": "home"}, _addr("home", "Berlin")]}
        )
        assert result["homeAddress"]["city"] == "Berlin"


class TestAuthoritativeUpdatePayload:
    def test_missing_phones_are_explicitly_cleared(self) -> None:
        result = to_graph_contact_update({"names": [{"displayName": "A"}]})
        assert result["homePhones"] == []
        assert result["businessPhones"] == []
        assert result["mobilePhone"] is None

    def test_missing_emails_cleared(self) -> None:
        assert to_graph_contact_update({})["emailAddresses"] == []

    def test_missing_birthday_cleared(self) -> None:
        assert to_graph_contact_update({})["birthday"] is None

    def test_missing_addresses_cleared_with_empty_subfields_not_null(self) -> None:
        # Graph answers "homeAddress": null with a bare 500 on PATCH.
        result = to_graph_contact_update({})
        for slot in ("homeAddress", "businessAddress", "otherAddress"):
            assert result[slot] == dict.fromkeys(ADDRESS_SUBFIELDS, "")

    def test_populated_address_carries_every_subfield(self) -> None:
        # Graph requires the full property set when updating an address, and a
        # subfield removed in Google must be cleared rather than left behind.
        result = to_graph_contact_update(
            {"addresses": [{"type": "home", "city": "Berlin"}]}
        )
        assert result["homeAddress"] == {
            "street": "",
            "city": "Berlin",
            "state": "",
            "postalCode": "",
            "countryOrRegion": "",
        }

    def test_contact_without_addresses_or_birthday(self) -> None:
        """The PATCH body that failed for every contact in production."""
        result = to_graph_contact_update(
            {
                "names": [{"displayName": "Jane Doe", "givenName": "Jane"}],
                "emailAddresses": [{"value": "jane@example.com"}],
            }
        )
        assert result["givenName"] == "Jane"
        assert result["emailAddresses"] == [
            {"address": "jane@example.com", "name": "Jane Doe"}
        ]
        assert result["birthday"] is None
        for slot in ("homeAddress", "businessAddress", "otherAddress"):
            assert result[slot] == dict.fromkeys(ADDRESS_SUBFIELDS, "")

    def test_address_change_changes_the_hash(self) -> None:
        before = to_graph_contact_update(
            {"addresses": [{"type": "home", "city": "Berlin", "streetAddress": "A 1"}]}
        )
        after = to_graph_contact_update(
            {"addresses": [{"type": "home", "city": "Berlin"}]}
        )
        assert contact_hash(before) != contact_hash(after)

    def test_missing_organization_cleared(self) -> None:
        result = to_graph_contact_update({})
        assert result["companyName"] is None
        assert result["jobTitle"] is None
        assert result["department"] is None

    def test_every_managed_field_is_present(self) -> None:
        result = to_graph_contact_update({})
        assert set(result) == set(MANAGED_GRAPH_FIELDS)

    def test_populated_values_survive(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact_update(person_full)
        assert result["givenName"] == "Jane"
        assert result["mobilePhone"] == "+49 170 1234567"
        assert result["homeAddress"]["city"] == "Berlin"
        assert result["businessAddress"] == dict.fromkeys(ADDRESS_SUBFIELDS, "")

    def test_create_payload_stays_sparse(self, person_full: dict[str, Any]) -> None:
        result = to_graph_contact(person_full)
        assert "businessPhones" not in result
        assert "businessAddress" not in result
        assert "otherAddress" not in result

    def test_empty_person_create_is_empty_but_update_clears_all(self) -> None:
        assert to_graph_contact({}) == {}
        update = to_graph_contact_update({})
        assert len(update) == len(MANAGED_GRAPH_FIELDS)
        empty_address = dict.fromkeys(ADDRESS_SUBFIELDS, "")
        assert all(v in ([], None, empty_address) for v in update.values())


class TestIsDeleted:
    def test_deleted_person(self, person_deleted: dict[str, Any]) -> None:
        assert is_deleted(person_deleted) is True

    def test_active_person(self, person_full: dict[str, Any]) -> None:
        assert is_deleted(person_full) is False

    def test_no_metadata_not_deleted(self) -> None:
        assert is_deleted({}) is False


class TestContactHash:
    def test_same_contact_same_hash(self, person_full: dict[str, Any]) -> None:
        contact = to_graph_contact(person_full)
        assert contact_hash(contact) == contact_hash(contact)

    def test_different_contacts_different_hash(
        self, person_full: dict[str, Any], person_no_year_birthday: dict[str, Any]
    ) -> None:
        h1 = contact_hash(to_graph_contact(person_full))
        h2 = contact_hash(to_graph_contact(person_no_year_birthday))
        assert h1 != h2

    def test_key_order_does_not_affect_hash(self) -> None:
        a = {"givenName": "A", "surname": "B"}
        b = {"surname": "B", "givenName": "A"}
        assert contact_hash(a) == contact_hash(b)

    def test_hash_is_64_char_hex(self, person_full: dict[str, Any]) -> None:
        h = contact_hash(to_graph_contact(person_full))
        assert len(h) == 64
        assert all(c in "0123456789abcdef" for c in h)
