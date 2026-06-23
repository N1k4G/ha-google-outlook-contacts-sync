"""Sync engine — orchestrates Google → Outlook contact synchronisation."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from typing import Any

from homeassistant.core import HomeAssistant
from homeassistant.util import dt as dt_util

from ..store import ContactSyncStore
from .dedup import (
    SELECT_FIELDS,
    OutlookContact,
    choose_keeper,
    find_duplicate_groups,
)
from .google_client import GoogleAuthError, GoogleContactsClient, SyncTokenExpiredError
from .graph_client import GraphContactsClient
from .mapping import contact_hash, is_deleted, to_graph_contact

_LOGGER = logging.getLogger(__name__)


class SyncAction(StrEnum):
    """The action the engine would take for a single Google person."""

    CREATE = "create"
    UPDATE = "update"
    DELETE = "delete"
    SKIP = "skip"


@dataclass
class _Decision:
    """Internal result of evaluating one person without applying anything."""

    action: SyncAction
    resource_name: str
    display_name: str
    graph_contact: dict[str, Any] | None = None
    existing_id: str | None = None
    content_hash: str | None = None


@dataclass
class PlannedChange:
    """A single contact change surfaced by a dry run."""

    resource_name: str
    display_name: str

    def as_dict(self) -> dict[str, str]:
        return {
            "resource_name": self.resource_name,
            "display_name": self.display_name,
        }


@dataclass
class SyncResult:
    """Outcome statistics from one sync run."""

    created: int = 0
    updated: int = 0
    deleted: int = 0
    skipped: int = 0
    failed: int = 0
    duplicates_removed: int = 0
    errors: list[str] = field(default_factory=list)
    duplicates: list[str] = field(default_factory=list)
    finished_at: datetime | None = None
    duration_seconds: float | None = None

    @property
    def changed(self) -> int:
        """Number of contacts actually mutated in Outlook."""
        return self.created + self.updated + self.deleted

    def summary(self) -> str:
        """Short human-readable status line."""
        parts = [
            f"{self.created} created",
            f"{self.updated} updated",
            f"{self.deleted} deleted",
        ]
        if self.duplicates_removed:
            parts.append(f"{self.duplicates_removed} dupes removed")
        if self.failed:
            parts.append(f"{self.failed} failed")
        return ", ".join(parts)

    def as_dict(self) -> dict[str, Any]:
        return {
            "created": self.created,
            "updated": self.updated,
            "deleted": self.deleted,
            "skipped": self.skipped,
            "failed": self.failed,
            "duplicates_removed": self.duplicates_removed,
            "errors": self.errors,
            "duplicates": self.duplicates,
            "finished_at": (
                self.finished_at.isoformat() if self.finished_at else None
            ),
            "duration_seconds": self.duration_seconds,
        }


@dataclass
class SyncPlan:
    """What a real sync *would* do — produced by a dry run, applies nothing."""

    to_create: list[PlannedChange] = field(default_factory=list)
    to_update: list[PlannedChange] = field(default_factory=list)
    to_delete: list[PlannedChange] = field(default_factory=list)
    finished_at: datetime | None = None

    @property
    def total(self) -> int:
        return len(self.to_create) + len(self.to_update) + len(self.to_delete)

    def as_dict(self) -> dict[str, Any]:
        return {
            "to_create": [c.as_dict() for c in self.to_create],
            "to_update": [c.as_dict() for c in self.to_update],
            "to_delete": [c.as_dict() for c in self.to_delete],
            "total": self.total,
            "finished_at": (
                self.finished_at.isoformat() if self.finished_at else None
            ),
        }


class SyncEngine:
    """Orchestrates one full or delta sync pass."""

    def __init__(
        self,
        hass: HomeAssistant,
        google_client: GoogleContactsClient,
        graph_client: GraphContactsClient,
        store: ContactSyncStore,
        delete_removed: bool = False,
        auto_remove_duplicates: bool = False,
    ) -> None:
        self._hass = hass
        self._google = google_client
        self._graph = graph_client
        self._store = store
        self._delete_removed = delete_removed
        self._auto_remove_duplicates = auto_remove_duplicates

    async def async_sync(self) -> SyncResult:
        """Delta sync pass: apply changed Google contacts to Outlook."""
        result = SyncResult()
        started = dt_util.utcnow()

        try:
            people, new_sync_token = await self._fetch_people(self._store.sync_token)
        except SyncTokenExpiredError:
            _LOGGER.warning("Google sync token expired; performing full re-sync")
            self._store.sync_token = None
            people, new_sync_token = await self._fetch_people(None)
        except GoogleAuthError:
            raise
        except Exception as exc:
            return self._fail_fetch(result, started, exc)

        for person in people:
            await self._process(self._decide(person), person, result)

        if self._auto_remove_duplicates:
            await self._dedupe(result)

        await self._finalize(result, started, new_sync_token)
        return result

    async def async_full_sync(self) -> SyncResult:
        """Full reconciliation: fetch all Google contacts, adopt matching
        Outlook contacts by name (so we don't create duplicates), and push any
        differences. Resets the delta token to the freshly returned one.
        """
        result = SyncResult()
        started = dt_util.utcnow()

        try:
            people, new_sync_token = await self._fetch_people(None)
        except GoogleAuthError:
            raise
        except Exception as exc:
            return self._fail_fetch(result, started, exc)

        adopt_index = await self._build_adopt_index()

        for person in people:
            await self._process(
                self._decide(person, adopt_index, force=True), person, result
            )

        if self._auto_remove_duplicates:
            await self._dedupe(result)

        # A full sync rebuilds state, so adopt the new token regardless of the
        # previous one (still only when nothing failed).
        self._store.sync_token = None
        await self._finalize(result, started, new_sync_token)
        _LOGGER.info("Full resync complete: %s", result.summary())
        return result

    async def async_plan(self) -> SyncPlan:
        """Compute what a delta sync would do, without applying or persisting."""
        plan = SyncPlan()

        try:
            people, _ = await self._fetch_people(self._store.sync_token)
        except SyncTokenExpiredError:
            # Do NOT clear the stored token here — that would affect real syncs.
            people, _ = await self._fetch_people(None)

        for person in people:
            decision = self._decide(person)
            change = PlannedChange(decision.resource_name, decision.display_name)
            if decision.action is SyncAction.CREATE:
                plan.to_create.append(change)
            elif decision.action is SyncAction.UPDATE:
                plan.to_update.append(change)
            elif decision.action is SyncAction.DELETE:
                plan.to_delete.append(change)

        plan.finished_at = dt_util.utcnow()
        _LOGGER.debug(
            "Dry run: %d create, %d update, %d delete",
            len(plan.to_create),
            len(plan.to_update),
            len(plan.to_delete),
        )
        return plan

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    async def _fetch_people(
        self, sync_token: str | None
    ) -> tuple[list[dict[str, Any]], str | None]:
        return await self._hass.async_add_executor_job(
            self._google.fetch_contacts, sync_token
        )

    def _fail_fetch(
        self, result: SyncResult, started: datetime, exc: Exception
    ) -> SyncResult:
        _LOGGER.error("Failed to fetch Google contacts: %s", exc)
        result.failed += 1
        result.errors.append(f"Google fetch failed: {exc}")
        result.finished_at = dt_util.utcnow()
        result.duration_seconds = (result.finished_at - started).total_seconds()
        return result

    async def _finalize(
        self, result: SyncResult, started: datetime, new_sync_token: str | None
    ) -> None:
        # Only advance the delta cursor when the run was clean — otherwise a
        # failed contact would be skipped forever on subsequent delta syncs.
        if new_sync_token and result.failed == 0:
            self._store.sync_token = new_sync_token
        elif new_sync_token and result.failed:
            _LOGGER.warning(
                "Keeping previous sync token: %d contact(s) failed this run",
                result.failed,
            )
        await self._store.async_save()
        result.finished_at = dt_util.utcnow()
        result.duration_seconds = (result.finished_at - started).total_seconds()

    async def _build_adopt_index(self) -> dict[str, list[str]]:
        """Map name key → unclaimed Outlook contact ids, for full-sync adoption."""
        raw = await self._hass.async_add_executor_job(
            self._graph.list_contacts_raw, SELECT_FIELDS
        )
        index: dict[str, list[str]] = {}
        for item in raw:
            contact = OutlookContact.from_graph(item)
            key = contact.name_key()
            if key and contact.contact_id not in self._store.mapped_ids():
                index.setdefault(key, []).append(contact.contact_id)
        return index

    def _decide(
        self,
        person: dict[str, Any],
        adopt_index: dict[str, list[str]] | None = None,
        force: bool = False,
    ) -> _Decision:
        """Determine the action for one person without performing it.

        When *adopt_index* is provided (full sync), an unmapped Google contact
        whose name matches an existing Outlook contact adopts that contact
        instead of creating a new one — preventing duplicates.

        When *force* is True the stored hash is ignored and existing contacts
        are always updated.  Used by full resync to recover from cases where
        the stored hash matches but Outlook is missing fields (e.g. birthday
        dropped by an earlier buggy implementation).
        """
        resource_name: str = person.get("resourceName", "")
        existing_id = self._store.lookup(resource_name)

        if is_deleted(person):
            if existing_id and self._delete_removed:
                return _Decision(
                    SyncAction.DELETE,
                    resource_name,
                    _display_name(None, resource_name),
                    existing_id=existing_id,
                )
            return _Decision(
                SyncAction.SKIP, resource_name, _display_name(None, resource_name)
            )

        graph_contact = to_graph_contact(person)
        display_name = _display_name(graph_contact, resource_name)
        if not graph_contact:
            return _Decision(SyncAction.SKIP, resource_name, display_name)

        current_hash = contact_hash(graph_contact)

        if existing_id is None and adopt_index is not None:
            existing_id = _claim_adopt(adopt_index, display_name)

        if existing_id is None:
            return _Decision(
                SyncAction.CREATE,
                resource_name,
                display_name,
                graph_contact=graph_contact,
                content_hash=current_hash,
            )

        if not force and self._store.get_hash(resource_name) == current_hash:
            return _Decision(SyncAction.SKIP, resource_name, display_name)

        return _Decision(
            SyncAction.UPDATE,
            resource_name,
            display_name,
            graph_contact=graph_contact,
            existing_id=existing_id,
            content_hash=current_hash,
        )

    async def _process(
        self, decision: _Decision, person: dict[str, Any], result: SyncResult
    ) -> None:
        """Execute a single decision, capturing per-contact failures."""
        resource_name = decision.resource_name
        try:
            await self._apply(decision, result)
        except Exception as exc:
            _LOGGER.error("Unexpected error processing %s: %s", resource_name, exc)
            result.failed += 1
            result.errors.append(f"{resource_name}: {exc}")

    async def _apply(self, decision: _Decision, result: SyncResult) -> None:
        if decision.action is SyncAction.CREATE:
            assert decision.graph_contact is not None
            contact_id = await self._hass.async_add_executor_job(
                self._graph.create_contact, decision.graph_contact
            )
            self._store.save_mapping(
                decision.resource_name, contact_id, decision.content_hash or ""
            )
            result.created += 1
        elif decision.action is SyncAction.UPDATE:
            assert decision.existing_id is not None
            assert decision.graph_contact is not None
            await self._hass.async_add_executor_job(
                self._graph.update_contact,
                decision.existing_id,
                decision.graph_contact,
            )
            self._store.save_mapping(
                decision.resource_name,
                decision.existing_id,
                decision.content_hash or "",
            )
            result.updated += 1
        elif decision.action is SyncAction.DELETE:
            assert decision.existing_id is not None
            await self._hass.async_add_executor_job(
                self._graph.delete_contact, decision.existing_id
            )
            self._store.delete_mapping(decision.resource_name)
            result.deleted += 1
        else:
            result.skipped += 1

    async def _dedupe(self, result: SyncResult) -> None:
        """Remove duplicate Outlook contacts (matched by name), keeping one."""
        try:
            raw = await self._hass.async_add_executor_job(
                self._graph.list_contacts_raw, SELECT_FIELDS
            )
        except Exception as exc:
            _LOGGER.error("Failed to list Outlook contacts for dedup: %s", exc)
            result.errors.append(f"Dedup listing failed: {exc}")
            return

        contacts = [OutlookContact.from_graph(item) for item in raw]
        groups = find_duplicate_groups(contacts)
        mapped_ids = self._store.mapped_ids()

        for group in groups:
            keeper = choose_keeper(group, mapped_ids)
            for contact in group.contacts:
                if contact.contact_id == keeper.contact_id:
                    continue
                try:
                    await self._hass.async_add_executor_job(
                        self._graph.delete_contact, contact.contact_id
                    )
                except Exception as exc:
                    _LOGGER.error(
                        "Failed to delete duplicate %s: %s", contact.contact_id, exc
                    )
                    result.errors.append(f"Dedup delete failed: {exc}")
                    continue
                self._store.repoint_mapping(contact.contact_id, keeper.contact_id)
                result.duplicates_removed += 1
                result.duplicates.append(group.key)
                _LOGGER.warning(
                    "Removed duplicate Outlook contact %s (name=%r); kept %s",
                    contact.contact_id,
                    group.key,
                    keeper.contact_id,
                )


def _claim_adopt(
    adopt_index: dict[str, list[str]], display_name: str
) -> str | None:
    """Pop and return an unclaimed Outlook contact id matching the name."""
    key = " ".join(display_name.split()).casefold()
    candidates = adopt_index.get(key)
    if candidates:
        return candidates.pop(0)
    return None


def _display_name(
    graph_contact: dict[str, Any] | None, resource_name: str
) -> str:
    """Best-effort human label for a contact change."""
    if graph_contact:
        if display := graph_contact.get("displayName"):
            return str(display)
        given = graph_contact.get("givenName", "")
        surname = graph_contact.get("surname", "")
        full = f"{given} {surname}".strip()
        if full:
            return full
    return resource_name or "(unknown)"
