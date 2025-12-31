
# SHAMROCK — Backend System Specification

**Document:** BACKEND.SPEC.md  
**Status:** Authoritative, Contractual  
**Audience:** Implementation AI Agent / Backend Developer  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

---

## 1. Purpose

This document defines the **backend architecture, authority model, write paths, and invariants** for SHAMROCK.

It exists to:
- Eliminate ambiguity during implementation
- Prevent logic drift between public sheets and backend data
- Guarantee auditability, safety, and long-term maintainability

This is **not** a tutorial. This is a **contract**.

---

## 2. Core Architecture

### Two-Tier Model

SHAMROCK uses a **strict separation** between:

1. **Public-Facing Sheets**
   - Human-readable
   - Editable in limited, controlled ways
   - Never authoritative

2. **Backend Spreadsheet (Private)**
   - Source of truth
   - Write-only by Apps Script
   - Never edited manually

All mutations originate from Apps Script and land in the backend.

---

## 3. Backend Spreadsheet

### Name (Recommended)

```
SHAMROCK — Backend (Source of Truth)
```

### Access Rules

- Owned by system administrators
- Not shared broadly
- Editors limited to service owners
- Apps Script has full access

---

## 4. Backend Tabs (Required)

### 4.1 `directory_backend`

Authoritative cadet records.

**Primary Key**
- `cadet_email` (University email)

**Fields**
- cadet_email
- last_name
- first_name
- as_year
- graduation_year
- flight
- squadron
- university
- phone
- dorm
- home_town
- home_state
- dob
- cip_broad
- cip_code
- afsc
- flight_path_status
- status
- photo_url
- notes
- created_at
- updated_at

Names are **display-only** and must never be used as join keys.

---

### 4.2 `events_backend`

Defines events independently of attendance.

**Primary Key**
- `event_id` (UUID or deterministic hash)

**Fields**
- event_id
- event_name
- event_type
- training_week
- event_date
- event_status (Draft | Published | Cancelled | Archived)
- affects_attendance (boolean)
- attendance_label
- expected_group
- flight_scope
- location
- notes
- created_at
- updated_at

Only `Published` events with `affects_attendance = true` may generate attendance rows.

---

### 4.3 `attendance_backend`

Authoritative attendance records.

**Composite Key**
- cadet_email
- event_id

**Fields**
- cadet_email
- event_id
- attendance_code
- source (form | menu | override | sync)
- updated_at

Attendance codes must match the canonical set.

---

### 4.4 `excusals_backend`

Private excusal workflow storage.

**Primary Key**
- `excusal_id`

**Fields**
- excusal_id
- cadet_email
- event_id
- request_timestamp
- reason (private)
- decision (Pending | Approved | Denied)
- decision_by
- decision_timestamp
- attendance_effect
- source

Only sanitized projections appear in public sheets.

---

### 4.5 `admin_actions`

Normalized queue of admin-triggered actions.

Used as a **command ledger**, not a database.

Fields:
- action_id
- actor_email
- action_type
- payload_json
- created_at
- processed_at
- status

---

### 4.6 `audit_log`

Immutable, append-only log.

**No edits. No deletes. Ever.**

Fields:
- timestamp
- actor_email
- action
- entity_type
- entity_id
- field
- old_value
- new_value
- source
- run_id

All mutations must produce exactly one audit entry.

---

## 5. Source-of-Truth Rules

| Data Type | Authority |
|---------|----------|
Cadet identity | directory_backend |
Event definitions | events_backend |
Attendance | attendance_backend |
Excusals | excusals_backend |
Admin actions | Apps Script |
Audit | audit_log |

Public sheets are **views and inputs only**.

---

## 6. Allowed Write Paths

### Valid Mutation Sources
- Apps Script menu actions
- Approved Google Forms
- Validated `onEdit` handlers (attendance cells only)

### Prohibited
- Manual backend edits
- Direct public-to-public writes
- Name-based joins
- Silent overwrites

---

## 7. Identity Model

- **Primary join key:** University email
- Names are cosmetic
- Email is immutable
- All historical data keyed by email

---

## 8. Attendance Codes (Canonical)

| Code | Meaning |
|----|--------|
P | Present |
E | Excused |
ES | Excused – Sport |
ER | Excusal Requested |
ED | Excusal Denied |
T | Tardy |
U | Unexcused |
UR | Unexcused – Report Submitted |
MU | Make-Up |
MRS | Medical / No PT |
N/A | Cancelled / Not Applicable |

Attendance percentage logic:
- Credit: P, E, ES, MU, MRS
- No Credit: U, ED
- Pending: ER
- Excluded: N/A

---

## 9. Admin Overrides

- Implemented via Apps Script menus
- No data-entry UI tables required
- All overrides must:
  1. Validate inputs
  2. Update backend
  3. Sync public views
  4. Write audit log entry

---

## 10. Events → Attendance Model

- Events exist independently
- Attendance rows created only when event is published
- Unpublishing does not delete attendance
- Cancelled events set attendance to N/A

---

## 11. Concurrency & Performance

- Use `LockService` on all writes
- Batch reads/writes
- No per-row API calls
- Idempotent operations
- Public sheets rebuilt from backend snapshots

---

## 12. Explicit Non-Goals

The backend must **not**:
- Store analytics
- Compute dashboards
- Depend on formatting
- Read chart data
- Trust public sheets blindly

---

## 13. Final Instruction

**Implement exactly as specified.**

This spec is contractual.
