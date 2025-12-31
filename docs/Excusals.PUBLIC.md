# SHAMROCK — Excusals (Public Specification)

**Document:** `Excusals.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Excusals`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

---

## Purpose

The `Excusals` tab provides a **public, sanitized view** of excusal requests and outcomes.

It allows cadets, cadre, and leadership to:
- See which excusals have been submitted
- Track approval / denial status
- Understand how an excusal affects attendance

This tab **does not** expose private justification text, attachments, or MFRs.

---

## Source of Truth

- Backend excusal records (restricted) are the authoritative source
- This sheet is a **projection** of those records
- Each row corresponds to **one cadet + one event**

Primary join key:
```
(cadet_email, event_id)
```

---

## Table Structure & Row Conventions

### Row 1 — Machine Header (Hidden, Protected)
- snake_case identifiers
- Used by Apps Script only
- Must never be edited or reordered

### Row 2 — Human Header (Visible)
- User-facing labels
- Forms the table header

### Row 3+ — Data Rows
- One row per excusal request
- Updated in-place when status changes

---

## Column Schema

### Row 1 — Machine Keys
```
request_id
event_id
cadet_email
cadet_last
cadet_first
cadet_flight
cadet_squadron
request_status
decision
decision_by
decision_timestamp
attendance_effect
submitted_timestamp
last_updated
public_notes
```

### Row 2 — Human Headers

| Column | Display Header |
|------|----------------|
| A | Request ID |
| B | Event |
| C | Email |
| D | Last Name |
| E | First Name |
| F | Flight |
| G | Squadron |
| H | Status |
| I | Decision |
| J | Decided By |
| K | Decided At |
| L | Attendance Effect |
| M | Submitted |
| N | Last Updated |
| O | Notes |

---

## Allowed Values

### Request Status
- Submitted (default when no decision is recorded)
- Approved (mirrors decision)
- Denied (mirrors decision)
- Withdrawn / Superseded (if the backend marks them explicitly)

---

### Decision
- Approved
- Denied
- (blank if pending)

---

### Attendance Effect
Defines how Attendance should be updated.

Allowed values:
- None / No Publish
- Set ER
- Set E
- Set ED

---

## Dummy Data Example

| Request ID | Event | Email | Last | First | Flight | Squadron | Status | Decision | Decided By | Decided At | Attendance | Submitted | Updated | Notes |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| EAR-2026-00041 | 2026S-TW12-MANDO | dhuggin2@nd.edu | Huggins | Declan | B | Blue | Approved | Approved | cadre@nd.edu | 2026-03-25 19:14 | Set E | 2026-03-24 22:01 | 2026-03-25 19:14 | — |
| EAR-2026-00058 | 2026S-TW12-LLAB | jdoe5@nd.edu | Doe | John | D | Gold | In Review | — | — | — | Set ER | 2026-03-24 21:33 | 2026-03-24 21:33 | — |
| EAR-2026-00060 | 2026S-TW11-SECONDARY | asmith3@nd.edu | Smith | Ava | A | Blue | Denied | Denied | c/commander@nd.edu | 2026-03-18 08:02 | Set ED | 2026-03-17 23:09 | 2026-03-18 08:02 | — |

---

## Privacy Rules

Public:
- Cadet identity
- Event
- Status and decision
- Attendance impact

Not public:
- Excusal reason
- MFR attachments
- Narrative justification
- Internal routing notes

---

## Attendance Interaction Rules

- Submitting an excusal may set `ER` if the event publishes to attendance
- Approval changes `ER → E`
- Denial changes `ER → ED`
- Existing attendance codes must not be overwritten improperly

---

## Governance

- Rows are system-managed
- Manual edits are discouraged
- All changes are logged in `Audit / Changelogs`

---

## Non-Goals

This sheet does not:
- Replace Attendance
- Store justification text
- Handle approvals directly
- Serve as an audit log

---

## Contractual Stability

Any changes to:
- Column order
- Allowed values
- Update semantics

must be reflected in this document and coordinated with backend logic.
