# SHAMROCK — Attendance (Public Specification)

**Document:** `Attendance.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Attendance`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

---

## Purpose

The `Attendance` tab is the **authoritative public record of attendance outcomes** for all SHAMROCK events.

It provides:
- A cadet-by-event attendance matrix
- A sanctioned surface for limited manual correction
- A clear, auditable view of presence, absence, and excusal state

The Attendance tab **does not define events**, **does not store excusal reasons**, and **does not contain historical audit data**.

---

## Source of Truth

Attendance operates under a **mediated source-of-truth model**:

- The **backend system state** is authoritative
- The Attendance sheet is an **approved write surface**
- All edits are validated and logged

There are **two sanctioned write paths**:
1. Automated ingestion (Forms, Excusals, Scripts)
2. Direct cell edits by authorized users

All writes converge to the same backend logic.

---

## Table Structure & Row Conventions

All SHAMROCK public tables follow the same structural contract:

### Row 1 — Machine Header (Hidden, Protected)
- Stable identifiers used by Apps Script
- Never shown to users
- Never edited manually
- Changing these requires code changes

### Row 2 — Human Header (Visible, Table Header)
- Human-readable column labels
- Forms the header row of the table

### Row 3+ — Data Rows
- One row per cadet
- Rows persist even if cadet becomes inactive

---

## Identity & Keys

### Primary Join Key
- **Email** (`cadet_email`)
  - Unique
  - Immutable
  - Hidden from users
  - Used to join against Directory, Events, and Excusals

### Display Names
- `Last Name` and `First Name` are display-only
- Names are never used as join keys
- Corrections must not affect email continuity

---

## Column Schema

### Row 1 — Machine Headers (hidden)

```
cadet_email
status
last_name
first_name
as_year
flight
squadron
<event_id_1>
<event_id_2>
...
```

Event IDs follow a stable format, e.g.:

```
2026S-TW01-LLAB
2026S-TW01-MANDO
2026S-TW01-SEC
```

---

### Row 2 — Human Headers (visible)

| Column | Display Header |
|------|----------------|
| A | Email *(hidden)* |
| B | Status *(hidden)* |
| C | Last Name |
| D | First Name |
| E | AS Year |
| F | Flight |
| G | Squadron |
| H+ | Event Labels (e.g., `TW-01 LLAB`) |

---

## Required Fields

The following fields must not be blank:
- Email
- Status
- Last Name
- First Name
- AS Year
- Flight
- Squadron

---

## Attendance Codes (Authoritative)

| Code | Meaning |
|----|--------|
| P | Present |
| E | Excused |
| ES | Excused – Sport |
| ER | Excusal Requested |
| ED | Excusal Denied |
| T | Tardy |
| U | Unexcused |
| UR | Unexcused – Report Submitted |
| MU | Make-up Complete |
| MRS | Medical / No PT |
| N/A | Cancelled / Not Applicable |

---

## Validation Rules

- One code per cell
- Codes must come from the approved list
- Blank cells are permitted temporarily
- Data validation dropdown is enforced

---

## Edit Rules & Permissions

### Who May Edit
- Cadre
- Flight Commanders
- Approved Administrators

### Publishing Behavior
- Attendance columns exist only for events where `affects_attendance = TRUE` and `status` is `Published`, `Cancelled`, or `Archived`.
- Cancelled events write `N/A` for all cadets.

### Edit Behavior
On any cell edit:
1. User authorization is verified
2. Cadet status is checked
3. Code validity is enforced against the canonical list
4. Backend state is updated
5. Entry is written to `Audit / Changelogs`
6. Dependent dashboards refresh

Inactive or Alumni cadets may be visually greyed out and optionally locked.

---

## Relationship to Other Tabs

| Tab | Dependency |
|----|-----------|
| Directory | Supplies roster and status |
| Events | Defines event columns |
| Excusals | Drives ER → E / ED transitions |
| Admin Overrides | Controlled mutation interface |
| Audit / Changelogs | Immutable edit history |

---

## Non-Goals

The Attendance tab does **not**:
- Define events or schedules
- Store excusal reasons
- Perform percentage calculations inline
- Delete historical attendance data

All such logic exists elsewhere in SHAMROCK.

---

## Change Control

Any change to:
- Attendance codes
- Column order
- Edit semantics
- Event ID format

Must be reflected in this document and reviewed before deployment.

This schema is considered **contractual** with the rest of SHAMROCK.
