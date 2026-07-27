"""Sync engine — orchestrates Google → Outlook contact synchronisation."""

from __future__ import annotations

import asyncio
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
    normalize_email,
)
from .google_client import GoogleContactsClient, SyncTokenExpiredError
from .graph_client import GraphContactsClient, MSAuthError
from .mapping import (
    contact_hash,
    is_deleted,
    to_graph_contact,
    to_graph_contact_update,
)

_LOGGER = logging.getLogger(__name__)


class SyncAction(StrEnum):
    """The action the engine would take for a single Google person."""

    CREATE = "create"
    UPDATE = "update"
    DELETE = "delete"
    SKIP = "skip"


@dataclass
class OutlookSnapshot:
    """One complete listing of Outlook contacts, taken by a full resync.

    Serves two purposes from a single Graph call: deciding which existing
    Outlook contact an unmapped Google contact may adopt, and telling whether a
    contact id we still have mapped exists at all.
    """

    all_contact_ids: set[str]
    unclaimed_by_email: dict[str, set[str]]

    @classmethod
    def build(
        cls, contacts: list[OutlookContact], mapped_ids: set[str]
    ) -> OutlookSnapshot:
        by_email: dict[str, set[str]] = {}
        for contact in contacts:
            if not contact.contact_id or contact.contact_id in mapped_ids:
                # Already owned by another Google contact — never a candidate.
                continue
            for email in contact.emails:
                by_email.setdefault(email, set()).add(contact.contact_id)
        return cls(
            all_contact_ids={c.contact_id for c in contacts if c.contact_id},
            unclaimed_by_email=by_email,
        )

    def claim_unique(self, emails: set[str], display_name: str) -> str | None:
        """Claim the one unmapped Outlook contact sharing an email, if unambiguous.

        Matching on name was the previous rule and is deliberately gone: two
        different people who share a name are common, and adopting on that basis
        overwrites a stranger's contact with someone else's data.
        """
        if not emails:
            return None

        candidates: set[str] = set()
        for email in emails:
            candidates |= self.unclaimed_by_email.get(email, set())

        if not candidates:
            return None
        if len(candidates) > 1:
            _LOGGER.warning(
                "Not adopting an existing Outlook contact for %r: %d unmapped "
                "contacts share its email address(es); creating a new contact "
                "instead",
                display_name,
                len(candidates),
            )
            return None

        claimed = candidates.pop()
        # Drop it from every email key, not just the one that matched, so it can
        # never be claimed a second time in this run.
        for ids in self.unclaimed_by_email.values():
            ids.discard(claimed)
        return claimed


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
            "finished_at": (self.finished_at.isoformat() if self.finished_at else None),
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
            "finished_at": (self.finished_at.isoformat() if self.finished_at else None),
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
        # Serialises every public run. The scheduled coordinator refresh, the
        # sync/full-sync services, both buttons and the dry run all reach this
        # one engine and the one store; without this they interleave at every
        # await and can double-create contacts or cross their sync tokens.
        self._run_lock = asyncio.Lock()

    async def async_sync(self) -> SyncResult:
        """Delta sync pass: apply changed Google contacts to Outlook."""
        async with self._run_lock:
            return await self._async_sync_unlocked()

    async def async_full_sync(self) -> SyncResult:
        """Full reconciliation of every Google contact against Outlook."""
        async with self._run_lock:
            return await self._async_full_sync_unlocked()

    async def async_plan(self) -> SyncPlan:
        """Compute what a sync would do, without applying or persisting.

        Takes the same lock as a real run: a preview computed from a store that
        is being mutated mid-run describes a state that never existed.
        """
        async with self._run_lock:
            return await self._async_plan_unlocked()

    # ------------------------------------------------------------------
    # Run implementations — these never take the lock themselves, so they can
    # call each other. asyncio.Lock is not reentrant.
    # ------------------------------------------------------------------

    async def _async_sync_unlocked(self) -> SyncResult:
        """Delta pass, or full reconciliation when there is no usable cursor.

        Without a valid sync token there is no delta to apply, and running the
        delta pipeline over a full fetch skips adoption, stale-mapping repair
        and drift correction — the very things a full fetch exists to do.
        """
        if self._store.sync_token is None:
            _LOGGER.debug("No Google sync token stored; running full reconciliation")
            return await self._async_full_sync_unlocked()

        result = SyncResult()
        started = dt_util.utcnow()

        try:
            people, new_sync_token = await self._fetch_people(self._store.sync_token)
        except SyncTokenExpiredError:
            _LOGGER.warning(
                "Google sync token expired; running full reconciliation instead"
            )
            self._store.sync_token = None
            return await self._async_full_sync_unlocked()

        for person in people:
            await self._process(self._decide(person), person, result)

        if self._auto_remove_duplicates:
            await self._dedupe(result)

        await self._finalize_delta(result, started, new_sync_token)
        return result

    async def _async_full_sync_unlocked(self) -> SyncResult:
        """Fetch every Google contact, adopt matching Outlook contacts (so we
        don't create duplicates), and push any differences.
        """
        result = SyncResult()
        started = dt_util.utcnow()

        people, new_sync_token = await self._fetch_people(None)

        snapshot = await self._build_snapshot()

        for person in people:
            self._drop_stale_mapping(person, snapshot)
            await self._process(
                self._decide(person, snapshot, force=True), person, result
            )

        if self._auto_remove_duplicates:
            await self._dedupe(result)

        await self._finalize_full(result, started, new_sync_token)
        _LOGGER.info("Full resync complete: %s", result.summary())
        return result

    async def _async_plan_unlocked(self) -> SyncPlan:
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

    async def _finalize_delta(
        self, result: SyncResult, started: datetime, new_sync_token: str | None
    ) -> None:
        """Advance the delta cursor only on a clean run.

        Keeping the previous token on failure means Google re-delivers the same
        changes next time, so a failed contact is retried instead of skipped
        forever.
        """
        if result.failed:
            _LOGGER.warning(
                "Keeping previous sync token: %d operation(s) failed this run",
                result.failed,
            )
        elif new_sync_token:
            self._store.sync_token = new_sync_token
        await self._persist(result, started)

    async def _finalize_full(
        self, result: SyncResult, started: datetime, new_sync_token: str | None
    ) -> None:
        """Set the fresh cursor, or drop it entirely when the run failed.

        The previous token is deliberately *not* restored on failure: a contact
        that failed reconciliation may have last changed before that cursor, so
        a later delta would never deliver it again. Clearing it routes the next
        run back into full reconciliation.
        """
        if result.failed:
            self._store.sync_token = None
            _LOGGER.warning(
                "Discarding sync token: %d operation(s) failed during full resync; "
                "the next run will reconcile in full again",
                result.failed,
            )
        else:
            self._store.sync_token = new_sync_token
        await self._persist(result, started)

    async def _persist(self, result: SyncResult, started: datetime) -> None:
        await self._store.async_save()
        result.finished_at = dt_util.utcnow()
        result.duration_seconds = (result.finished_at - started).total_seconds()

    async def _build_snapshot(self) -> OutlookSnapshot:
        """Load every Outlook contact once, for adoption and existence checks.

        A failure here propagates: acting on a partial listing would create
        duplicates for contacts that merely failed to load.
        """
        raw = await self._hass.async_add_executor_job(
            self._graph.list_contacts_raw, SELECT_FIELDS
        )
        contacts = [OutlookContact.from_graph(item) for item in raw]
        return OutlookSnapshot.build(contacts, self._store.mapped_ids())

    def _drop_stale_mapping(
        self, person: dict[str, Any], snapshot: OutlookSnapshot
    ) -> None:
        """Forget a mapping whose Outlook contact no longer exists.

        Without this, a contact the user deleted in Outlook is patched by id on
        every full resync, fails every time, and is never recreated. Afterwards
        the Google contact is simply unmapped, so it goes through the normal
        adoption-or-create path.

        Only safe during a full resync, where the snapshot is a complete and
        verified listing — a delta sync must not pay for per-contact existence
        checks, and a partial listing would delete live mappings.
        """
        resource_name: str = person.get("resourceName", "")
        mapped_id = self._store.lookup(resource_name)
        if mapped_id is not None and mapped_id not in snapshot.all_contact_ids:
            _LOGGER.warning(
                "Outlook contact %s mapped to %s no longer exists; unlinking so it "
                "can be recreated",
                mapped_id,
                resource_name,
            )
            self._store.delete_mapping(resource_name)

    def _decide(
        self,
        person: dict[str, Any],
        snapshot: OutlookSnapshot | None = None,
        force: bool = False,
    ) -> _Decision:
        """Determine the action for one person without performing it.

        When *snapshot* is provided (full resync), an unmapped Google contact
        may adopt an existing Outlook contact instead of creating a duplicate —
        but only on an unambiguous email match, never on name alone.

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

        create_contact = to_graph_contact(person)
        update_contact = to_graph_contact_update(person)
        display_name = _display_name(create_contact, resource_name)

        # Change detection runs on the authoritative state so that removing the
        # last value of a field still registers as a change.
        current_hash = contact_hash(update_contact)

        if existing_id is None and snapshot is not None:
            existing_id = snapshot.claim_unique(
                _emails_of(create_contact), display_name
            )

        if existing_id is None:
            if not create_contact:
                # Nothing mappable and nothing in Outlook yet — nothing to do.
                return _Decision(SyncAction.SKIP, resource_name, display_name)
            return _Decision(
                SyncAction.CREATE,
                resource_name,
                display_name,
                graph_contact=create_contact,
                content_hash=current_hash,
            )

        if not force and self._store.get_hash(resource_name) == current_hash:
            return _Decision(SyncAction.SKIP, resource_name, display_name)

        # A mapped contact is updated even when every supported Google field is
        # now empty — that is exactly the case that must clear Outlook.
        return _Decision(
            SyncAction.UPDATE,
            resource_name,
            display_name,
            graph_contact=update_contact,
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
        except MSAuthError:
            raise
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


def _emails_of(graph_contact: dict[str, Any]) -> set[str]:
    """Normalized email addresses of a mapped Graph contact body."""
    entries = graph_contact.get("emailAddresses") or []
    return {
        normalized
        for entry in entries
        if (normalized := normalize_email(str(entry.get("address") or "")))
    }


def _display_name(graph_contact: dict[str, Any] | None, resource_name: str) -> str:
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
