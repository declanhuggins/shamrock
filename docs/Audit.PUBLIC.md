# SHAMROCK — Audit / Changelogs (Public Specification)

**Document:** `Audit.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Audit / Changelogs`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

---

## Purpose

`Audit / Changelogs` is the **append-only system log** for SHAMROCK.

It exists to:
- Make all state changes explainable
- Provide a forensic trail (who changed what, when, and why)
- Support rollback (optional) and accountability

This tab is the **single shared log** for:
- Directory updates (new cadets, status changes, corrections)
- Attendance edits (manual overrides, back-propagation outcomes)
- Event changes (publish/unpublish, date edits, cancellations)
- Excusal workflow (submit, approve/deny, withdraw)
- Automation actions (rebuild roster, re-sync columns, triggers)

---

## Source of Truth

- The log itself is authoritative for **what happened**
- It is written only by Apps Script
- It is never edited in place

If rollback is implemented later, rollback actions also append audit rows.

---

## Table Structure & Row Conventions

### Row 1 — Machine Header (Hidden, Protected)
- Stable snake_case column identifiers
- Used by Apps Script
- Never edited manually
- Never reordered

### Row 2 — Human Header (Visible)
- Display labels for humans
- Defines the table header

### Row 3+ — Append-Only Entries
- New entries only (no row reuse)
- Old entries never mutated

---

## Governance & Protection

- Entire tab is protected:
  - No edits by normal users
  - No edits by admins (except via script)
- The only allowed manual interaction is:
  - filtering / sorting
  - copying out data for reports

---

## Column Schema

### Row 1 — Machine Keys
```
audit_id
timestamp
actor_email
actor_role
action
target_sheet
target_table
target_key
target_range
event_id
request_id
old_value
new_value
result
reason
notes
source
script_version
run_id
```

### Row 2 — Human Headers

| Column | Display Header |
|---|---|
| A | Audit ID |
| B | Timestamp |
| C | Actor Email |
| D | Actor Role |
| E | Action |
| F | Target Sheet |
| G | Target Table |
| H | Target Key |
| I | Target Range |
| J | Event ID |
| K | Request ID |
| L | Old Value |
| M | New Value |
| N | Result |
| O | Reason |
| P | Notes |
| Q | Source |
| R | Script Version |
| S | Run ID |

---

## Field Semantics

### audit_id
- Unique, stable identifier for the log row
- Recommended format:
  - `AUD-YYYYMMDD-HHMMSS-<random>` (e.g., `AUD-20260325-191402-A7K3`)

### timestamp
- ISO-like string in script timezone, e.g. `2026-03-25 19:14:02`
- Always populated

### actor_email
- `Session.getActiveUser().getEmail()` when available
- Fallback: `"unknown"` (should be rare)

### actor_role
- Derived from allowlist / role mapping (backend config)
- Examples: `Cadet`, `Flight Commander`, `Cadre`, `Admin`, `System`

### action
A short verb phrase describing the operation (stable enums preferred), e.g.:
- `directory.add_cadet`
- `directory.update_field`
- `attendance.set_code`
- `attendance.clear_code`
- `attendance.backprop_from_public`
- `events.create`
- `events.update`
- `events.publish_toggle`
- `excusals.submit`
- `excusals.approve`
- `excusals.deny`
- `system.sync_roster`
- `system.rebuild_attendance_columns`

### target_sheet / target_table
- `target_sheet`: the spreadsheet tab name (e.g., `Attendance`)
- `target_table`: logical table name (usually equals the tab, but can differ if a tab has multiple tables)

### target_key
- The primary identifier for the record affected
- Prefer stable keys:
  - Directory / Attendance: `cadet_email`
  - Events: `event_id`
  - Excusals: `request_id` and/or `(cadet_email,event_id)`

### target_range
- A1 notation when the change is cell/range-specific, e.g. `Attendance!H14`
- Empty for non-cell changes (e.g. “rebuild columns”)

### event_id / request_id
- Optional, populated when relevant

### old_value / new_value
- String snapshots (best effort)
- For multi-cell edits, either:
  - store a compact JSON string, or
  - store a summary and include details in `notes`
- Keep under reasonable size (Apps Script cell limits apply)

### result
Allowed values:
- `success`
- `rejected`
- `skipped`
- `error`

### reason
- Short explanation of why the change was made
- Required for admin overrides and excusal decisions
- Optional for automated sync actions

### notes
- Extra structured info (compact)
- Can include validation failures, row numbers, counts, etc.

### source
Allowed values:
- `menu`
- `form`
- `public_edit`
- `trigger`
- `system`

### script_version
- Git SHA or tag (recommended) injected at build time if feasible
- Otherwise a manually maintained version string

### run_id
- Unique per execution (UUID recommended)
- Groups multiple audit entries from one operation

---

## Dummy Data Example

| Audit ID | Timestamp | Actor Email | Role | Action | Target Sheet | Target Table | Target Key | Target Range | Event ID | Request ID | Old | New | Result | Reason | Notes | Source | Version | Run ID |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| AUD-20260324-220112-K9F2 | 2026-03-24 22:01:12 | dhuggin2@nd.edu | Cadet | excusals.submit | Excusals | Excusals | dhuggin2@nd.edu |  | 2026S-TW12-MANDO | EAR-2026-00041 |  | Submitted | success |  | created request row | form | v0.1.0 | RUN-6c2a |
| AUD-20260324-220115-X2P1 | 2026-03-24 22:01:15 | system | System | attendance.set_code | Attendance | Attendance | dhuggin2@nd.edu | Attendance!N14 | 2026S-TW12-MANDO | EAR-2026-00041 |  | ER | success | request submitted | protectExisting=true | trigger | v0.1.0 | RUN-6c2a |
| AUD-20260325-191402-A7K3 | 2026-03-25 19:14:02 | cadre@nd.edu | Cadre | excusals.approve | Excusals | Excusals | dhuggin2@nd.edu |  | 2026S-TW12-MANDO | EAR-2026-00041 | In Review | Approved | success | medical appointment | decision recorded | menu | v0.1.1 | RUN-91bd |
| AUD-20260325-191410-T4M8 | 2026-03-25 19:14:10 | cadre@nd.edu | Cadre | attendance.set_code | Attendance | Attendance | dhuggin2@nd.edu | Attendance!N14 | 2026S-TW12-MANDO | EAR-2026-00041 | ER | E | success | approved excusal | ER→E only | menu | v0.1.1 | RUN-91bd |
| AUD-20260326-081233-R0Q7 | 2026-03-26 08:12:33 | c/lead@nd.edu | Admin | attendance.backprop_from_public | Attendance | Attendance | jdoe5@nd.edu | Attendance!P22 | 2026S-TW12-LLAB |  | A | P | success | roster discrepancy fix | change mirrored to backend | public_edit | v0.1.1 | RUN-3aa1 |
| AUD-20260328-000001-SYS1 | 2026-03-28 00:00:01 | system | System | system.sync_roster | Directory | Directory |  |  |  |  |  |  | success | daily sync | rows=214 updated=3 added=1 | trigger | v0.1.1 | RUN-7f0c |

Notes:
- `system` is allowed for automated triggers
- `reason` may be blank for fully automated maintenance actions
- `old_value/new_value` may be blank when “create row” actions occur

---

## Required Logging Rules (Contract)

The system must append an audit row when any of these happen:
- A directory row is created, removed (soft), or changed
- Any attendance cell is changed (including back-propagation)
- An event is created/edited/cancelled/publish toggled
- Any excusal request is created or changes status/decision
- Any automation sync runs (at least one summary row per run)

---

## Non-Goals

This sheet does not:
- Replace backend state
- Store private excusal reasons (beyond a short `reason` string if approved for public)
- Serve as a UI for making changes

It is strictly a log.

---

## Change Control

Any change to:
- column order
- action names
- allowed values
- required logging rules

must be reflected in this document and coordinated with backend implementation.
