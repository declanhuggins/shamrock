# SHAMROCK — Events (Public Specification)

**Document:** `Events.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Events`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

This document defines the **authoritative schema, constraints, and operating rules** for the public-facing **Events** sheet.

---

## Purpose

The `Events` tab is the **source of truth for event definitions** in SHAMROCK.

It provides:
- A canonical list of events for the term
- Stable identifiers that map 1:1 to Attendance columns (when published)
- Scheduling and classification metadata (TW, type, scope, expected group)
- A controlled lifecycle for events (draft → published → cancelled/archived)

The Events tab **does not**:
- Store attendance outcomes (that is `Attendance`)
- Store excusal reasons or approvals (that is `Excusals`)
- Store audit history (that is `Audit / Changelogs`)

---

## Source of Truth and Publishing Model

- The `Events` sheet is the **source of truth** for events.
- `Attendance` is the **source of truth** for attendance outcomes.
- Publishing is **one-way** from Events → Attendance:
  - Events define what columns exist in Attendance.
  - Attendance does not create events.

**Critical rules:**
- Each *published* event produces exactly one Attendance column **when `affects_attendance = TRUE`**.
- Events with `affects_attendance = FALSE` never create Attendance columns, regardless of status.

---

## Table Structure & Row Conventions

All SHAMROCK public tables follow this pattern:

### Row 1 — Machine Header (Hidden, Protected)
- Stable `snake_case` column identifiers
- Used exclusively by Apps Script
- Never edited manually
- Never reordered without code changes

### Row 2 — Human Header (Visible, Table Header)
- Display labels for users
- Forms the header row of the table

### Row 3+ — Data Rows
- One row per event
- Rows are append-only for historical integrity

---

## Identity & Keys

### Primary System Key
- `event_id` (unique, immutable once published)

The `event_id` is:
- The join key used across Attendance, Excusals, and Audit logs
- The machine header used in Attendance row 1 (for published events)

**Event IDs must never be edited** after an event is published.

---

## Column Schema (Authoritative Order)

### Row 1 — Machine Keys (hidden)

```
event_id
term
training_week
event_type
display_label
attendance_label
expected_group
flight_scope
status
start_datetime
end_datetime
location
notes
created_at
created_by
```

### Row 2 — Human Headers (visible)

| Column | Display Header |
|------|----------------|
| A | Event ID |
| B | Term |
| C | Training Week |
| D | Event Type |
| E | Display Name |
| F | Attendance Column Label |
| G | Expected Group |
| H | Flight Scope |
| I | Status |
| J | Start Date/Time |
| K | End Date/Time |
| L | Location |
| M | Notes |
| N | Created At |
| O | Created By |

---

## Required Fields

The following fields must not be blank:
- Event ID
- Term
- Training Week
- Event Type
- Display Name
- Status

Fields required **when published to Attendance**:
- Attendance Column Label

---

## Allowed Values & Validation Rules

### Term
Format: `YYYY{S|F}` (examples: `2026S`, `2026F`)

---

### Training Week
Allowed values:
- `TW-01` … `TW-16`

---

### Event Type
Allowed values:
- `LLAB`
- `Mando`
- `Secondary`
- `Other`

`Other` is permitted for non-attendance items, administrative deadlines, reminders, etc.

---

### Flight Scope
Allowed values:
- `All`
- `A`
- `B`
- `C`
- `D`
- `E`
- `F`

---

### Expected Group
Suggested allowed values (can be enumerated later if you want strict DV):
- `All Cadets`
- `AS100/150`
- `AS200/250`
- `AS300`
- `AS400`
- `Optional`

---

### Status (Lifecycle + Publishing Behavior)

Allowed values (matches backend + code):

- **Draft**  
  Event is defined but not yet published to attendance.

- **Published**  
  Event is active and published to Attendance (if `affects_attendance = TRUE`).

- **Cancelled**  
  Event will not occur. If already published, Attendance is kept and set to `N/A`.

- **Archived**  
  Historical event; locked from edits (except notes). Attendance column remains.

### Status → Attendance Publishing Rules

| Status | Attendance Column? | Notes |
|---|---:|---|
| Draft | No | Not published |
| Published | Yes (if `affects_attendance = TRUE`) | One column per event_id |
| Cancelled | Yes (if previously published) | Column remains and is set to `N/A` |
| Archived | Yes (if previously published) | Column remains, locked |

---

## Event ID Format

Event IDs must be globally unique and stable.

Recommended format (matches current implementation):
- `{term}-{training_week}-{type}`

Examples:
- `2026S-TW01-LLAB`
- `2026S-TW01-MANDO`
- `2026S-TW01-SEC`

If you later need multiple Secondary events in one TW, add a suffix:
- `2026S-TW01-SEC-01`, `2026S-TW01-SEC-02`

---

## Attendance Column Label Rules

- Used for the **human header** in Attendance (row 2)
- Can be renamed without breaking joins, as long as `event_id` remains stable
- Recommended labels:
  - `TW-01 LLAB`
  - `TW-01 Mando`
  - `TW-01 Secondary`

---

## Editing Rules & Governance

- `event_id` is immutable after publish.
- Do not delete event rows (append-only).
- Status changes are permitted but must follow the publishing rules above.
- If an event moves to `Published`, it must have an `attendance_label` when `affects_attendance = TRUE`.

---

## System Dependencies

The Events table feeds:
- **Attendance** — column creation + labeling (published events only)
- **Excusals** — matching excusal requests to an event definition (via TW + type)
- **Dashboards** — event counts, calendar views, weekly summaries
- **Audit / Changelogs** — event lifecycle edits

---

## Dummy Example (Schema-accurate)

Row 1 (hidden machine header):
```
event_id | term | training_week | event_type | display_label | attendance_label | expected_group | flight_scope | status | start_datetime | end_datetime | location | notes | created_at | created_by
```

Sample rows:
- `2026S-TW01-LLAB` | `2026S` | `TW-01` | `LLAB` | `LLAB` | `TW-01 LLAB` | `All Cadets` | `All` | `Published` | `2026-01-15 1500` | `2026-01-15 1700` | `Jordan Hall` | `Weekly LLAB` | `2025-12-01 0930` | `System`
- `2026S-TW01-MANDO` | `2026S` | `TW-01` | `Mando` | `Mando PT` | `TW-01 Mando` | `All Cadets` | `All` | `Published` | `2026-01-16 0600` | `2026-01-16 0700` | `Rockne` | `Morning PT` | `2025-12-01 0930` | `System`
- `2026S-TW01-SEC` | `2026S` | `TW-01` | `Secondary` | `Leadership Talk` | `TW-01 Secondary` | `Optional` | `All` | `Draft` | `2026-01-17 1900` | `2026-01-17 2000` | `Hesburgh` | `Tracked, not in attendance` | `2025-12-01 0930` | `System`
- `2026S-TW02-MANDO` | `2026S` | `TW-02` | `Mando` | `Mando PT` | `TW-02 Mando` | `All Cadets` | `All` | `Cancelled` | `2026-01-23 0600` | `2026-01-23 0700` | `Rockne` | `Weather cancellation` | `2025-12-01 0930` | `System`

---

## Change Control

Any change to:
- `event_id` format
- Status semantics
- Publishing rules
- Column order or meaning

must be reflected in this document and reviewed before deployment.

The Events schema is contractual with Attendance and Excusals.
