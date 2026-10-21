"""Maps Google People API Person objects to Microsoft Graph Contact dicts."""

from __future__ import annotations

import hashlib
import json
from datetime import date
from typing import Any

from ..const import BIRTHDAY_PLACEHOLDER_YEAR

# The Graph contact properties this integration owns. Everything else on an
# Outlook contact is left untouched.
#
# These three tuples are the single source of truth for building payloads,
# clearing removed values and hashing for change detection — keeping them apart
# is what let sparse payloads and stored hashes drift out of sync.
MANAGED_COLLECTION_FIELDS = (
    "emailAddresses",
    "homePhones",
    "businessPhones",
)
MANAGED_SCALAR_FIELDS = (
    "givenName",
    "surname",
    "displayName",
    "fileAs",
    "mobilePhone",
    "birthday",
    "companyName",
    "jobTitle",
    "department",
)
MANAGED_ADDRESS_FIELDS = (
    "homeAddress",
    "businessAddress",
    "otherAddress",
)

MANAGED_GRAPH_FIELDS = (
    MANAGED_COLLECTION_FIELDS + MANAGED_SCALAR_FIELDS + MANAGED_ADDRESS_FIELDS
)

# Every property of Graph's physicalAddress. Graph requires the full set when
# updating an address, and an address cannot be cleared with null: PATCHing
# "homeAddress": null fails with a bare 500 instead of a 400.
ADDRESS_SUBFIELDS = (
    "street",
    "city",
    "state",
    "postalCode",
    "countryOrRegion",
)

# Graph phone type buckets
_PHONE_HOME = {"home"}
_PHONE_MOBILE = {"mobile"}
_PHONE_WORK = {"work", "work_fax", "company_main"}

# Google address type → Graph address slot. Anything unrecognised (including
# "other" and user-defined labels) belongs in otherAddress, never business.
_ADDRESS_TYPE_MAP: dict[str, str] = {
    "home": "homeAddress",
    "work": "businessAddress",
}
_DEFAULT_ADDRESS_FIELD = "otherAddress"


def to_graph_contact(person: dict[str, Any]) -> dict[str, Any]:
    """Convert a Google Person to a *sparse* Graph Contact body, for creation.

    Only populated fields are included. Suitable for POST, where absent
    properties simply stay unset — use :func:`to_graph_contact_update` for
    PATCH, which must also clear fields removed in Google.
    """
    contact: dict[str, Any] = {}

    name = _primary_or_first(person.get("names", []))
    _map_names(name, contact)
    _map_emails(person, contact, _display_name_from(name))
    _map_phones(person, contact)
    _map_birthday(person, contact)
    _map_addresses(person, contact)
    _map_organization(person, contact)

    return contact


def to_graph_contact_update(person: dict[str, Any]) -> dict[str, Any]:
    """Convert a Google Person to an *authoritative* Graph Contact body.

    Every managed property is present, with an explicit empty value where Google
    has no data: ``[]`` for lists, ``None`` for scalars and an address with every
    subfield ``""`` for addresses. Graph treats properties omitted from a PATCH
    as unchanged, so a sparse body can never clear a field the user deleted in
    Google — the value would survive in Outlook forever while the stored hash
    claimed it was synced.
    """
    contact = to_graph_contact(person)

    for graph_field in MANAGED_COLLECTION_FIELDS:
        contact.setdefault(graph_field, [])
    for graph_field in MANAGED_SCALAR_FIELDS:
        contact.setdefault(graph_field, None)
    for graph_field in MANAGED_ADDRESS_FIELDS:
        # An empty address is every subfield set to "", never null; a populated
        # one carries all subfields too, so a removed street does not survive.
        address = contact.get(graph_field, {})
        contact[graph_field] = {sub: address.get(sub, "") for sub in ADDRESS_SUBFIELDS}

    return contact


def contact_hash(graph_contact: dict[str, Any]) -> str:
    """Return a stable SHA-256 digest of a Graph contact dict for change detection.

    Callers must hash the authoritative representation, not the sparse one:
    hashing a sparse body makes "field removed in Google" indistinguishable from
    "field never set".
    """
    serialized = json.dumps(graph_contact, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(serialized.encode()).hexdigest()


def is_deleted(person: dict[str, Any]) -> bool:
    """Return True if the Google People API marks this person as deleted."""
    return bool(person.get("metadata", {}).get("deleted", False))


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _primary_or_first(items: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not items:
        return None
    for item in items:
        if item.get("metadata", {}).get("primary"):
            return item
    return items[0]


def _display_name_from(name: dict[str, Any] | None) -> str:
    """Best available human name for the contact, or ""."""
    if name is None:
        return ""
    if display := name.get("displayName", ""):
        return str(display)
    given = name.get("givenName", "")
    family = name.get("familyName", "")
    return f"{given} {family}".strip()


def _map_names(name: dict[str, Any] | None, contact: dict[str, Any]) -> None:
    if name is None:
        return
    given = name.get("givenName", "")
    family = name.get("familyName", "")
    display = name.get("displayName", "")

    if given:
        contact["givenName"] = given
    if family:
        contact["surname"] = family
    if display:
        contact["displayName"] = display
        contact["fileAs"] = display


def _map_emails(
    person: dict[str, Any], contact: dict[str, Any], display_name: str
) -> None:
    emails: list[dict[str, Any]] = person.get("emailAddresses", [])
    if not emails:
        return

    graph_emails: list[dict[str, str]] = []
    for email in emails:
        address = email.get("value", "").strip()
        if not address:
            continue
        # Graph's emailAddresses[].name is the display name for the address, not
        # a type label — Google's home/work/other has no Graph counterpart here
        # and must not be smuggled into it.
        entry: dict[str, str] = {"address": address}
        if display_name:
            entry["name"] = display_name
        graph_emails.append(entry)

    if graph_emails:
        contact["emailAddresses"] = graph_emails


def _map_phones(person: dict[str, Any], contact: dict[str, Any]) -> None:
    phones: list[dict[str, Any]] = person.get("phoneNumbers", [])
    if not phones:
        return

    home_phones: list[str] = []
    mobile_phone: str | None = None
    business_phones: list[str] = []

    for phone in phones:
        number = phone.get("value", "").strip()
        if not number:
            continue
        google_type = phone.get("type", "other").lower()
        if google_type in _PHONE_MOBILE:
            if mobile_phone is None:
                mobile_phone = number
        elif google_type in _PHONE_WORK:
            business_phones.append(number)
        else:
            home_phones.append(number)

    if home_phones:
        contact["homePhones"] = home_phones
    if mobile_phone is not None:
        contact["mobilePhone"] = mobile_phone
    if business_phones:
        contact["businessPhones"] = business_phones


def _map_birthday(person: dict[str, Any], contact: dict[str, Any]) -> None:
    birthdays: list[dict[str, Any]] = person.get("birthdays", [])
    bday = _primary_or_first(birthdays)
    if bday is None:
        return

    bday_date: dict[str, Any] = bday.get("date", {})
    month: int | None = bday_date.get("month")
    day: int | None = bday_date.get("day")
    year: int | None = bday_date.get("year")  # None or 0 means no year

    if month is None or day is None:
        return

    effective_year = year if year else BIRTHDAY_PLACEHOLDER_YEAR

    try:
        d = date(effective_year, month, day)
    except ValueError:
        return

    # Graph Contact birthday must be an ISO 8601 primitive string (DateTimeOffset).
    # Sending a DateTimeTimeZone object causes a 400 "PrimitiveValue expected" error.
    contact["birthday"] = f"{d.isoformat()}T00:00:00Z"


def _map_addresses(person: dict[str, Any], contact: dict[str, Any]) -> None:
    """Fill each Graph address slot with the first valid Google address of that type.

    An address with no transferable fields is skipped without consuming its slot,
    so a later valid address of the same type still lands.
    """
    addresses: list[dict[str, Any]] = person.get("addresses", [])
    if not addresses:
        return

    for addr in addresses:
        graph_addr: dict[str, str] = {}
        if street := addr.get("streetAddress", ""):
            graph_addr["street"] = street
        if city := addr.get("city", ""):
            graph_addr["city"] = city
        if region := addr.get("region", ""):
            graph_addr["state"] = region
        if postal := addr.get("postalCode", ""):
            graph_addr["postalCode"] = postal
        if country := addr.get("countryCode", "") or addr.get("country", ""):
            graph_addr["countryOrRegion"] = country

        if not graph_addr:
            continue

        google_type = addr.get("type", "other").lower()
        graph_field = _ADDRESS_TYPE_MAP.get(google_type, _DEFAULT_ADDRESS_FIELD)
        if graph_field not in contact:
            contact[graph_field] = graph_addr


def _map_organization(person: dict[str, Any], contact: dict[str, Any]) -> None:
    orgs: list[dict[str, Any]] = person.get("organizations", [])
    org = _primary_or_first(orgs)
    if org is None:
        return

    if name := org.get("name", ""):
        contact["companyName"] = name
    if title := org.get("title", ""):
        contact["jobTitle"] = title
    if dept := org.get("department", ""):
        contact["department"] = dept
