# SHAMROCK — Directory (Public Specification)

**Document:** `Directory.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Directory`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

This document defines the **authoritative schema, constraints, and operating rules** for the public-facing **Directory** sheet.

---

## Purpose

The `Directory` tab is the **canonical public roster** for SHAMROCK.

It provides:
- Human-readable cadet identity and grouping
- Structural data for Flights, Squadrons, Universities, and AFROTC status
- The primary lookup source for Attendance, Events, Excusals, Dashboards, and Forms

The Directory **does not** track:
- Attendance outcomes
- Event definitions
- Excusal workflow state
- Audit history

Those concerns are handled in dedicated tabs.

---

## Table Structure & Row Conventions

All SHAMROCK public tables follow this pattern:

### Row 1 — Machine Header (Hidden, Protected)
- Stable, snake_case column identifiers
- Used exclusively by Apps Script
- **Never edited manually**
- **Never reordered without code changes**

### Row 2 — Human Header (Visible)
- Display labels for users
- Forms the header row of the table
- May be renamed cosmetically *only if Row 1 remains unchanged*

### Row 3+ — Data Rows
- One row per cadet
- No duplicate identities

---

## Identity & Keys

### Primary System Key
- **Email** (`cadet_email`)
  - Unique
  - Immutable once assigned
  - Used for all joins across sheets

### Name Rules
- **Required:** `First Name`, `Last Name`
- Names are **display-only**
- Names are **never used as join keys**
- Any name corrections must not break email continuity

---

## Column Schema (Authoritative Order)

### Row 1 — Machine Keys (hidden)

```
last_name
first_name
as_year
graduation_year
flight
squadron
university
cadet_email
phone
dorm
home_town
home_state
dob
cip_broad
cip_code
afsc
flight_path_status
status
photo_url
notes
updated_at
source
```

### Row 2 — Human Headers (visible)

| Column | Display Header |
|------|----------------|
| A | Last Name |
| B | First Name |
| C | AS Year |
| D | Class Year |
| E | Flight |
| F | Squadron |
| G | University |
| H | Email |
| I | Phone |
| J | Dorm |
| K | Home Town |
| L | Home State |
| M | DOB |
| N | CIP Broad Area |
| O | CIP Code |
| P | Desired / Assigned AFSC |
| Q | Flight Path Status |
| R | Status |
| S | Photo Link |
| T | Notes |
| U | Last Updated |
| V | Source |

---

## Required Fields

The following fields **must not be blank**:
- Last Name
- First Name
- AS Year
- Flight
- Squadron
- University
- Email
- Status

Rows missing any required field are considered **invalid**.

---

## Allowed Values & Validation Rules

### AS Year
Allowed values:
- AS100
- AS150
- AS200
- AS250
- AS300
- AS400

---

### University
Allowed values:
- Notre Dame
- Holy Cross
- St. Mary's
- Trine
- Valparaiso

---

### Squadron
Allowed values:
- Blue
- Gold

---

### Flight
Allowed values:
- A
- B
- C
- D
- E
- F

---

### Dorm
Validation rule:
- If `University = Notre Dame` → must be a valid Notre Dame dorm
- Else → must be exactly `Cross-Town`

---

### Flight Path Status
Allowed values (exact text):
- Participating 1/4
- Enrolled 2/4
- Active 3/4
- Ready 4/4

---

### Status
Allowed values:
- Active
- Leave
- Inactive
- Alumni

Inactive or Alumni cadets may be visually de-emphasized by downstream logic but **must remain in the table** for historical integrity.

---

## Editing Rules & Governance

- The Directory is **primarily protected**
- Typical users should **not directly edit** this tab
- Approved admins may edit:
  - Flight
  - Squadron
  - Status
  - Contact information
- New cadets are introduced via:
  - Approved Forms
  - Admin Overrides
- Existing cadets **must never be duplicated**
  - Updates must occur in-place by email match

---

## System-Managed Columns

The following columns are **write-only by Apps Script**:
- `Last Updated`
- `Source`

These may be hidden from normal users.

---

## System Dependencies

The Directory feeds:
- **Attendance** — roster generation & matching
- **Events** — participant resolution
- **Excusals** — identity resolution & email routing
- **Forms** — name lists and validation
- **Dashboards** — aggregated analytics

Birthday calculations, attendance percentages, and analytics **do not live here**.

---

## Non-Goals

The Directory does **not**:
- Track attendance per event
- Store excusal outcomes
- Define events or schedules
- Store audit logs

Each concern is handled in its own dedicated tab.

---

## Change Control

Any change to:
- Column order
- Column meaning
- Allowed values
- Identity rules

**must** be reflected in this document and reviewed before deployment.

The Directory schema is considered **contractual** with the rest of SHAMROCK.
