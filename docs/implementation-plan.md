# Implementation plan — open issue backlog (#6–#33)

> **Status: implemented.** All 9 stages are on `feat/sync-reliability-backlog`
> as 12 commits. This document is kept as the review guide for the PR; the
> "Deviations" section at the end records where the build differed from the plan.

Branch: `feat/sync-reliability-backlog` → **one PR**, 9 ordered commit stages.

28 open issues delivered as a single pull request. The staging below is not a
PR split — it is the commit order inside this branch. Each stage must leave the
test suite green on its own so a regression can still be bisected to a stage.

**Do not squash on merge.** Squashing collapses nine independently-verified
stages into one commit and throws away the only bisection handle this PR has.

Working-tree note: the CI workflow files already carry uncommitted
`actions/checkout@v4 → v6` / `setup-python@v5 → v6` bumps. They ride along and
land in stage 9, which touches the same files.

---

## Review guidance for a PR this size

This PR touches nearly every file in the integration. Two stages are
destructive-path changes and deserve the bulk of review attention:

- **Stage 5** (#25, #30, #18) — rewrites which existing Outlook contact a Google
  contact adopts, and adds orphan deletion. A bug here overwrites or deletes
  contacts the user never wanted touched.
- **Stage 6** (#27, #31b) — the auto-dedup delete path.

Everything else is either additive, corrective, or docs. Review stages 5 and 6
commit-by-commit; the rest can be reviewed as a diff.

---

## Dependency graph (why this order)

```
S1 graph client (#6 #10 #26)  ──┐
S2 store (#28, #31a)          ──┤
S3 mapping (#29 #32 #24 #9)   ──┼──> S5 full-resync reconciliation (#25 #30 #18)
S4 engine lock + cursor (#8 #33 #14) ┘        │
                                              ├──> S6 dedup safety (#27 #31b)
                                              └──> S7 lifecycle (#16 #17 #19 #20 #7)
S8 metadata/docs (#11 #21 #15 #12 #23)   (independent — can land any time)
S9 tests + CI gate (#22)                 (last — measures everything above)
```

Hard couplings found while reading the code:

| Coupling | Reason |
|---|---|
| #10 + #26 | Both must raise the *same* `GraphRequestError`. Splitting them creates two parallel error classes, which #26 explicitly forbids. |
| #24 + #29 + #32 | All three rewrite `mapping.py`. #32 adds `otherAddress`, which #24's authoritative field list must include, and #29 changes the same `_map_emails()` body. |
| #32 + #9 | `_COMPLETENESS_FIELDS` and `SELECT_FIELDS` must move together, and `otherAddress` must be added to both. |
| #8 + #33 | #33 needs `_async_full_sync_unlocked()`, which only exists after #8 splits public/private methods. Doing #33 first would deadlock the non-reentrant lock. |
| #25 + #30 | Both consume the single full Outlook listing. #30 needs `all_contact_ids` from the same snapshot #25 builds the adoption index from. |
| #19 → #25, #8 | The initial background run routes into full-resync; it must not run against name-only adoption or race a user button press. |
| #31 | Splits across two layers: the `save_mapping()` invariant belongs in stage 2, the dedup group guard in stage 6. |

---

## Stage 1 — Graph client hardening (#6, #10, #26)

**Files:** `sync/graph_client.py`, `tests/test_graph_client.py` (new)

`create_contact` currently calls `.get()` on a `requests.Response`
([graph_client.py:72](../custom_components/google_outlook_contacts_sync/sync/graph_client.py#L72)),
so every create raises `AttributeError`, is swallowed by `_process`, and blocks
the sync token forever. This is the highest-impact bug in the backlog — it goes
first so the rest of the branch is built on a client that can actually write.

1. Add `GraphRequestError(Exception)` — the single non-auth Graph failure class
   shared by #10 and #26.
2. Add `_raise_request_failure(operation: str) -> Never`: calls `_require_auth()`
   first (raises `MSAuthError` when the account went unauthenticated), otherwise
   raises `GraphRequestError`.
3. `create_contact`: parse `result.json()`, read `id` from the body, raise
   `GraphRequestError` on a missing/empty id. Never return `""`.
4. `update_contact` / `delete_contact`: falsy response → `_raise_request_failure`.
5. `list_contacts_raw`: replace the `break` with `_raise_request_failure`; never
   return a partial page set. Raise on non-dict JSON body, on `.json()` throwing,
   and when `value` is not a list.
6. **#13 lifted forward from stage 9:** `html.escape()` in `_error_page`, and
   defensively in `_success_page`. Two lines, unauthenticated endpoint, no reason
   to make it wait for the test stage. Its regression test still lands in stage 9.

**Tests:** response-shaped mocks with `.json()` only — a `MagicMock` that answers
both `.get()` and `.json()` would hide the original #6 bug. Cover: POST with id,
POST without id, failed POST/PATCH/DELETE, paging over `@odata.nextLink` (second
request must not re-send `$select`/`$top`), failure on page 1 and page 2, auth
invalid vs. auth valid, invalid JSON body.

---

## Stage 2 — Store integrity (#28, #31 store half)

**Files:** `store.py`, `tests/test_store.py` (new)

`_DEFAULT_DATA` is a module-level nested mutable dict copied shallowly
([store.py:32](../custom_components/google_outlook_contacts_sync/store.py#L32)),
so every store instance in a process shares the same `mapping`/`hashes` objects.

1. Replace `_DEFAULT_DATA` with `_new_default_data()` factory; use it in
   `__init__` and in `async_load`.
2. On load, normalize defensively: `dict(stored.get("mapping", {}))`,
   `dict(stored.get("hashes", {}))`, `sync_token` as `str | None`. No schema
   migration, no `STORAGE_VERSION` bump — the on-disk format is unchanged.
3. Add `resource_names_for_contact_id(contact_id) -> set[str]`.
4. `save_mapping()` raises `MappingConflictError` when `contact_id` is already
   mapped to a *different* `resource_name`. Re-saving the same pair stays legal.
5. Add `async_remove()` thin wrapper over `Store.async_remove()` — consumed by
   #18 in stage 5.

**Tests:** two fresh stores are independent; a later-created store inherits
nothing; hashes isolated; loaded data not aliased to the mock dict; save writes
instance-local state; `save_mapping` conflict rejected and idempotent re-save
allowed; existing many-to-one conflict detectable.

---

## Stage 3 — Mapping correctness (#29, #32, #24, #9)

**Files:** `sync/mapping.py`, `sync/dedup.py`, `sync/engine.py`,
`tests/test_mapping.py`, `tests/test_engine.py`

Today `to_graph_contact()` only ever emits sparse payloads. Since Graph PATCH
treats absent properties as unchanged, deleting a field in Google updates the
stored hash but never clears the Outlook value — permanent silent divergence.

1. Introduce one central definition of the managed Graph properties in
   `mapping.py` so payload building, clearing and hashing cannot drift.
2. Split the API:
   - `to_graph_contact(person)` — sparse, for CREATE (unchanged behaviour).
   - `to_graph_contact_update(person)` — authoritative: every managed property
     present, with `[]` for collections and `None` for nullable scalars and
     addresses.
3. `_map_emails`: drop `_EMAIL_TYPE_MAP`. `emailAddresses[].name` is Graph's
   *display name*, not a type. Emit `{"address": ...}` plus `"name": display_name`
   only when a real contact name exists. Reuse the display name already derived
   for `displayName`; don't re-parse the person.
4. `_map_addresses`: three buckets — `home → homeAddress`, `work → businessAddress`,
   `other`/unknown → `otherAddress`. Keep "first valid address per slot"; stop
   mutating the input list via `pop()`.
5. `dedup.py`: add `otherAddress` to `_COMPLETENESS_FIELDS`, and expand
   `SELECT_FIELDS` to every counted property (`homePhones`, `businessPhones`,
   `mobilePhone`, `homeAddress`, `businessAddress`, `otherAddress`, `companyName`,
   `jobTitle`, `birthday`). Right now 8 of 12 counted fields are never requested,
   so the destructive keeper heuristic effectively ignores them.
6. `engine._decide`: hash the authoritative state, not the sparse payload. A
   mapped contact whose Google data is now entirely empty must produce UPDATE,
   not SKIP.
7. `engine._apply`: CREATE sends the sparse body, UPDATE sends the authoritative
   body. Hash is stored only after a successful write.

**Migration note (call out in CHANGELOG):** changing the hash basis invalidates
every stored hash, so the first sync after upgrade re-PATCHes all mapped
contacts once. That is intended and self-correcting.

**Verify before merge:** confirm Graph accepts `"birthday": null` for clearing.
If a specific property needs a different clear value, adjust *only* that
property, comment why, and cover it with a test — do not fall back to omitting it.

---

## Stage 4 — Engine concurrency + cursor state machine (#8, #33, #14)

**Files:** `sync/engine.py`, `tests/test_engine.py`

Three entry points (`async_sync`, `async_full_sync`, `async_plan`) reach the same
engine and store through two different coordinators with no mutual exclusion.

1. Add an instance-local `asyncio.Lock` in `SyncEngine.__init__` (not module-level).
2. Split every public method into a locked wrapper plus
   `_async_{sync,full_sync,plan}_unlocked()`. This is what makes #33's internal
   routing possible without self-deadlock. Waiting callers wait — never
   `if lock.locked(): return`.
3. **#33 routing:** `_async_sync_unlocked()` delegates to
   `_async_full_sync_unlocked()` when `store.sync_token is None`, and again on
   `SyncTokenExpiredError` (after clearing the token). Today both cases do a full
   *fetch* through the *delta* pipeline — no snapshot, no adoption, no `force=True`,
   no orphan reconciliation — while logging "performing full re-sync".
4. **Cursor semantics** — split or parameterize `_finalize()`:
   - delta ok → store new token; delta failed → keep the previous token;
   - full ok → store the fresh token; full failed → persist `sync_token = None`
     so the next scheduled run retries the real full pipeline.
   Remove the current `self._store.sync_token = None` before `_finalize()` in
   `async_full_sync`, which makes the "Keeping previous sync token" log a lie.
5. **#14:** delete `_fail_fetch`. Let fetch errors propagate; the coordinator
   already converts them to `UpdateFailed`, which keeps `last_update_success`
   false and stops `LastSyncSensor` from jumping to the failure timestamp.

**Tests:** deterministic `asyncio.Event`/future gating only — no `sleep()`.
Delta↔full, delta↔dry-run, full↔dry-run serialization; lock released after an
exception; no double CREATE from overlapping runs; tokenless and expired both
enter the full pipeline; full-failure persists `None`; delta-failure keeps the
old token; no deadlock when `async_sync` routes into the full pipeline.

---

## Stage 5 — Full-resync reconciliation (#25, #30, #18) ⚠️ destructive

**Files:** `sync/engine.py`, `sync/dedup.py`, `store.py`, `__init__.py`,
`README.md`, `tests/test_engine.py`

The most data-destructive area in the branch. Adoption currently claims the first
Outlook contact with a matching normalized name
([engine.py:445](../custom_components/google_outlook_contacts_sync/sync/engine.py#L445)),
so an unrelated same-name contact gets overwritten with Google data on every
full resync.

Split this stage into at least three commits (adoption, stale-ID repair, orphan
reconciliation) so a regression in one does not require reverting all three.

1. Replace `_build_adopt_index()` with a single `OutlookSnapshot` built from one
   Graph listing, exposing both `all_contact_ids: set[str]` and an email-keyed
   adoption index. No second full listing.
2. **#25 adoption rule:** match only on normalized email (`strip()` + `casefold()`).
   Adopt only when exactly one unmapped Outlook contact shares an email. Zero or
   multiple candidates → CREATE (and log a warning for the ambiguous case). Name
   alone never adopts. A successful claim removes that id from *every* index
   entry, so one Outlook id can serve at most one Google contact per run.
3. **#30 stale mapping repair:** in the full pipeline only, if
   `store.lookup(rn)` returns an id absent from `all_contact_ids`, drop the stale
   mapping and treat the contact as unmapped (then #25 adoption or CREATE). No
   "PATCH fails → CREATE" fallback: not every PATCH failure means the contact is
   gone. Delta sync does no per-contact existence checks.
4. **#18 orphan reconciliation:** add a `UNLINK` action —
   `store.delete_mapping()`, no Graph write, no `result.deleted`. Delta delete
   with `delete_removed=False` unlinks; with `True` it deletes then unmaps.
   For full resync, compute `active_resource_names` from the complete Google
   result *before* per-contact processing (a contact with no mappable fields is
   still active), then reconcile `store.all_resource_names() - active`.
   A failed remote delete keeps the mapping for retry, increments `failed`, and
   does not increment `deleted`.
5. Orphan reconciliation runs **only** after a fully successful Google fetch, and
   mapping deletions run **only** against a fully successful Outlook snapshot.
   Stage 1 guarantees a partial listing can no longer masquerade as success.
6. Add `async_remove_entry()` to delete the store file on entry removal.
   `async_unload_entry()` must keep the data — reload and restart depend on it.
7. README: replace "adopts existing Outlook contacts by name" with the unique-email
   rule, in both the Full resync section and the `full_sync` service description
   in all `strings.json`/translation files.

**Tests:** the six adoption cases from #25, the five stale-mapping cases from
#30, and the eight reconciliation cases from #18 — including that a stale mapping
removed before a failed CREATE stays removed.

---

## Stage 6 — Dedup safety and error handling (#27, #31 dedup half) ⚠️ destructive

**Files:** `sync/engine.py`, `tests/test_engine.py`, `tests/test_dedup.py`,
`README.md`

1. **#31 group guard:** before `choose_keeper()`, count *distinct currently mapped*
   Outlook ids in the group. More than one → skip the whole group with a warning,
   no delete, no repoint. Checking `len(group.contacts)` is not sufficient.
   Otherwise two Google resources end up pointing at one Outlook id and overwrite
   each other forever.
2. **#27 error handling:** re-raise `MSAuthError` from both the dedup listing and
   each dedup delete so the coordinator's reauth chain still fires. Non-auth
   failures increment `result.failed` (not just `result.errors`) so the sync token
   is withheld and `SyncProblemBinarySensor` turns on. Processing continues with
   the remaining groups.
3. `repoint_mapping()` and `duplicates_removed` only after a *successful* delete.
4. Record `result.duplicates` once per successfully handled group instead of once
   per deleted contact.
5. README duplicate-handling section: groups containing more than one tracked
   contact are never auto-merged.

---

## Stage 7 — Config entry lifecycle (#16, #17, #19, #20, #7)

**Files:** `__init__.py`, `coordinator.py`, `config_flow.py`,
`sync/google_client.py`, `const.py`, tests

1. **#16:** `GoogleContactsClient.from_token_dict` uses `.get()` instead of
   `token_dict["client_id"]`. In `async_setup_entry`, an empty token or a missing
   `client_id`/`client_secret`/`refresh_token` sets `DATA_REAUTH_PROVIDER` to
   `"google"` and raises `ConfigEntryAuthFailed` — matching the Microsoft path
   instead of crashing setup with `KeyError`.
2. **#20:** register both callback views exactly once behind a
   `views_registered` guard in `hass.data[DOMAIN]`, via one shared helper called
   from `async_setup` and from both config-flow steps.
3. **#17:** register the services once (first entry) and remove them on unload of
   the last entry, resolving the coordinator from `hass.data` at call time rather
   than closing over it. Unify refresh semantics: the `sync_now` service and the
   Sync-now button both use `async_request_refresh()`.
4. **#19:** stop awaiting `async_config_entry_first_refresh()` during setup.
   Order becomes: build clients → load store → build engine/coordinators → publish
   runtime data → `async_forward_entry_setups` → update listener → services →
   `entry.async_create_background_task(...)` for exactly one initial run. Choose
   `async_full_resync()` when `store.sync_token is None`, else
   `async_request_refresh()`. A failing initial run surfaces through the
   coordinator; it must not retroactively fail the completed setup.
5. **#7 token persistence:** one shared `_async_persist_tokens()` comparing
   `ms_backend.get_token_dict()` and `google_client.get_token_dict()` against
   `entry.data`, writing **one** `async_update_entry` when either changed, and
   nothing when neither did. Run it in a `finally` around every engine run
   (delta, full, dry-run — a dry run can still refresh Google credentials), inside
   the stage 4 lock boundary. A persistence failure must never mask the original
   sync or auth error; log it separately.
6. **Reload decoupling:** store `DATA_OPTIONS_SNAPSHOT = dict(entry.options)` at
   setup; the update listener compares the snapshot and returns without reloading
   when options are unchanged. Without this, token persistence triggers a reload,
   which starts a sync, which refreshes a token — a loop. No "ignore next update"
   boolean; it is race-prone under rapid updates.

**Tests:** MS-only change, Google-only change, both changed (exactly one
`async_update_entry`), nothing changed (no write), token persisted despite a later
sync failure, data update causes no reload, options change causes exactly one
reload, dry-run persists a Google refresh; setup returns while the initial sync is
still blocked on a future; tokenless → full, tokened → delta, exactly one initial
run; services gone after unload; views registered once across repeated flow starts.

---

## Stage 8 — Metadata, translations, docs (#11, #21, #15, #12, #23)

**Files:** `hacs.json`, `manifest.json`, `strings.json`, `translations/*.json`,
`README.md`, `store.py`

Independent of every other stage — land it early if the branch needs an easy
green commit.

1. **#11:** raise `hacs.json` `homeassistant` to `2024.12.0`. The code already
   requires `ConfigFlowResult` (2024.4), `DataUpdateCoordinator(config_entry=...)`
   (~2024.11/12), and stage 8 adds `single_config_entry` (2024.11). Today HACS
   would happily install onto 2024.1 and fail at import. Document the floor in
   README + CHANGELOG.
2. **#21:** drop `"quality_scale": "custom"` (not a valid HA value) and add
   `"single_config_entry": true`. Keep the manual `_async_current_entries()` check
   as a harmless fallback.
3. **#15:** add the missing `error.ms_url_failed` and `abort.missing_external_url`
   to `translations/en.json` — both are used by `config_flow.py` and exist only in
   `strings.json`/`de.json`, so English users see raw keys. Correct
   "Google Cloud OAuth 2.0 Desktop client" → "Web application client" in every
   locale; the code and README both require a web client with a redirect URI.
   Add a key-parity test between `strings.json` and each `translations/*.json`.
4. **#12:** HA `.storage/` is plain JSON protected only by filesystem permissions.
   Fix the claim in `README.md` and the `ContactSyncStore` docstring, and note that
   config-folder backups contain OAuth refresh tokens in clear text.
5. **#23:** add a Trademarks & Legal section — unofficial, not affiliated with or
   endorsed by Google LLC or Microsoft Corporation; each user operates their own
   OAuth credentials and owns compliance with the Google API Services User Data
   Policy.

---

## Stage 9 — Test coverage and CI gate (#22)

**Files:** `tests/` (new files), `.github/workflows/lint.yaml`,
`requirements_test.txt`, `pyproject.toml`

**Correction to the issue text** (worth posting as a comment on #22, but the
issue stays open — the rest of it is unbuilt): `lint.yaml` **already** runs
`ruff` (via `astral-sh/ruff-action`) and `mypy` on push and PR. The genuine gaps
are `ruff format --check`, unpinned tool versions, the coverage floor, and the
complete absence of tests for four modules.

1. Add `ruff format --check .` to `lint.yaml` rather than duplicating ruff/mypy
   into `test.yaml`. Pin `ruff` and `mypy` with version bounds in
   `requirements_test.txt` so CI does not depend on transitive luck.
2. New `tests/test_google_client.py`: first full fetch with `requestSyncToken`,
   delta fetch with `syncToken`, `nextPageToken` paging, `nextSyncToken` taken
   from the last response, HTTP 410 → `SyncTokenExpiredError`, HTTP 400 +
   `EXPIRED_SYNC_TOKEN` → same, any other `HttpError` *not* misclassified,
   `RefreshError` → `GoogleAuthError`, token serialization after refresh.
3. New `tests/test_http.py`: valid state resumes the flow, unknown/missing state →
   400, state consumed once (replay fails), OAuth error path, HTML escaping
   (regression test for the **#13** fix shipped in stage 1 — both views are
   unauthenticated with `requires_auth = False`, so `error_description` is
   reflected XSS today), both success paths.
4. Fill remaining lifecycle gaps not already covered by stages 1–8.
5. Raise `fail_under` from 50 → **80** only after measuring actual coverage with
   the new tests in place. This is the last commit on the branch.

---

## Summary

| Stage | Issues | Risk | Notes |
|---|---|---|---|
| 1 | #6, #10, #26, #13 | Low | Unblocks contact creation entirely — must be first |
| 2 | #28, #31a | Low | Foundation for stages 5/6 |
| 3 | #29, #32, #24, #9 | Medium | One-time re-PATCH of all contacts after upgrade |
| 4 | #8, #33, #14 | Medium | Structural split of the engine |
| 5 | #25, #30, #18 | **High** | Destructive adoption/delete paths — split into ≥3 commits |
| 6 | #27, #31b | **High** | Destructive dedup path |
| 7 | #16, #17, #19, #20, #7 | Medium | Reload-loop risk if options/data are not separated |
| 8 | #11, #21, #15, #12, #23 | Low | Docs/metadata only, order-independent |
| 9 | #22 | Low | Coverage gate goes to 80 % last |

All 28 issues close with this PR. Reference them in the PR body as
`Closes #6, Closes #7, …` so GitHub closes them on merge.

Open questions to settle during implementation, not before:

- Does Graph accept `null` for every managed property being cleared (stage 3)?
- Should `sync_now` keep `async_refresh()` semantics for users who script it, or
  move to `async_request_refresh()` as #17 asks (stage 7)? The issue is explicit,
  so the plan follows it, but it does change debounce behaviour for existing
  automations.


---

## Deviations from this plan

Recorded during implementation. Nothing was dropped; all 28 issues are closed.

1. **#13 (XSS escaping) moved from stage 9 to stage 1.** Two lines on an
   unauthenticated endpoint; with everything in one PR there was no reason to
   hold it behind the test stage. Its regression test still landed in stage 9.

2. **Part of #12 landed in stage 2.** The false "encrypted" claim in the
   `ContactSyncStore` docstring was corrected while that class was being
   rewritten. The README half stayed in stage 8.

3. **A prep commit was added before stage 1.** `tests/conftest.py` read JSON
   fixtures with the platform default encoding, so the suite failed on Windows
   and passed on Linux CI. Fixed first to make local verification possible.

4. **Stage 9's CI change is smaller than #22 asked for.** The issue says CI only
   ran pytest; `lint.yaml` already ran ruff and mypy. Duplicating them into
   `test.yaml` would have added two redundant jobs. What was actually missing —
   `ruff format --check`, pinned tool versions, a meaningful coverage floor —
   was added instead. This correction is worth posting as a comment on #22.

5. **Stage 9 gained entity, coordinator and diagnostics tests** not in the
   original scope. Coverage after the planned work was 73%, short of the 80%
   gate the issue requires; the coordinator's exception-to-reauth mapping was
   the largest untested behaviour worth fixing.

6. **The engine test harness changed shape.** `_make_engine` gained
   `stored_sync_token`, defaulting to a real cursor. After #33, a `None` cursor
   routes into full reconciliation, so every pre-existing delta test would have
   silently become a full-sync test.

## Verification performed

Each stage's regression tests were run against the *unfixed* code to confirm
they fail. Reintroducing:

| Defect | Tests that fail |
|---|---|
| `.get()` on the Graph response (#6) | 5 |
| shallow-copied store defaults (#28) | 3 |
| sparse-only update payload (#24) | 2 |
| no engine lock (#8) | 4 |
| no stale-mapping repair (#30) | 3 |
| old dedup error handling (#27/#31) | 5 |
| `html.escape` removed (#13) | 3 |
| `ms_url_failed` removed from en.json (#15) | 2 |

Final state: 238 tests, 84% coverage, `ruff check`, `ruff format --check` and
`mypy` clean.

## Known gaps

- **Graph null-clearing is unverified against a live tenant.** #24 sends `None`
  for managed scalar and address properties being cleared. They are documented
  as nullable, but this has not been exercised against a real mailbox. Editing
  one Google contact to remove a phone number would confirm it.
- **`config_flow.py` (24%) and `auth/ms_auth.py` (35%) remain lightly tested.**
  Both are OAuth plumbing that only a live provider exercises end to end.
- **The HA test plugin cannot run on Windows** — `homeassistant.runner` imports
  `fcntl`. The suite runs with `-p no:homeassistant` locally; CI on Linux runs it
  unrestricted. Lifecycle tests therefore mock the HA boundary rather than using
  the `hass` fixture.
