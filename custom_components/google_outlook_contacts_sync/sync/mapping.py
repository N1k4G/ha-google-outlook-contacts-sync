"""Maps Google People API Person objects to Microsoft Graph Contact dicts."""

from __future__ import annotations

import hashlib
import json
from datetime import date
from typing import Any

from ..const import BIRTHDAY_PLACEHOLDER_YEAR

# Graph email address type labels
_EMAIL_TYPE_MAP: dict[str, str] = {
    "home": "personal",
    "work": "work",
    "other": "other",
}

# Graph phone type buckets
_PHONE_HOME = {"home"}
_PHONE_MOBILE = {"mobile"}
_PHONE_WORK = {"work", "work_fax", "company_main"}


def to_graph_contact(person: dict[str, Any]) -> dict[str, Any]:
    """Convert a Google People API Person to a Graph Contact request body.

    Only fields that are populated in the Google person are included —
    missing fields are left absent so Graph doesn't overwrite existing data
    with empty values on partial updates.
    """
    contact: dict[str, Any] = {}

    _map_names(person, contact)
    _map_emails(person, contact)
    _map_phones(person, contact)
    _map_birthday(person, contact)
    _map_addresses(person, contact)
    _map_organization(person, contact)

    return contact


def contact_hash(graph_contact: dict[str, Any]) -> str:
    """Return a stable SHA-256 digest of a Graph contact dict for change detection."""
    serialized = json.dumps(graph_contact, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(serialized.encode()).hexdigest()


def is_deleted(person: dict[str, Any]) -> bool:
    """Return True if the Google People API marks this person as deleted."""
    return bool(
        person.get("metadata", {}).get("deleted", False)
    )


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


def _map_names(person: dict[str, Any], contact: dict[str, Any]) -> None:
    names: list[dict[str, Any]] = person.get("names", [])
    name = _primary_or_first(names)
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


def _map_emails(person: dict[str, Any], contact: dict[str, Any]) -> None:
    emails: list[dict[str, Any]] = person.get("emailAddresses", [])
    if not emails:
        return

    graph_emails: list[dict[str, str]] = []
    for email in emails:
        address = email.get("value", "").strip()
        if not address:
            continue
        google_type = email.get("type", "other").lower()
        graph_type = _EMAIL_TYPE_MAP.get(google_type, "other")
        graph_emails.append({"address": address, "name": graph_type})

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
    addresses: list[dict[str, Any]] = person.get("addresses", [])
    if not addresses:
        return

    graph_addresses: list[dict[str, str]] = []
    for addr in addresses:
        google_type = addr.get("type", "other").lower()
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

        graph_addr["type"] = "home" if google_type == "home" else "business"
        graph_addresses.append(graph_addr)

    if graph_addresses:
        _assign_addresses(graph_addresses, contact)


def _assign_addresses(
    graph_addresses: list[dict[str, str]], contact: dict[str, Any]
) -> None:
    """Write first home and first business address into the contact dict."""
    contact.pop("homeAddress", None)
    contact.pop("businessAddress", None)

    for addr in graph_addresses:
        addr_type = addr.pop("type", "home")
        if addr_type == "home" and "homeAddress" not in contact:
            contact["homeAddress"] = addr
        elif addr_type == "business" and "businessAddress" not in contact:
            contact["businessAddress"] = addr


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
