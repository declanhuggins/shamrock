# SHAMROCK — Cadre & Leadership (Public Specification)

**Document:** `CadreLeadership.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Cadre & Leadership`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

This document defines the **schema, intent, and governance** for the public-facing **Cadre & Leadership** sheet.

This tab serves as the **authoritative human directory** for cadre, detachment leadership, and key cadet leaders.

---

## Purpose

The `Cadre & Leadership` tab exists to provide:
- A clear chain of command
- Contact information for leadership roles
- A single, trusted reference for “who do I contact?”

It is intentionally **simple, manual, and stable**.

This tab is **not** intended for automation-heavy workflows.

---

## Automation Rules

- ❌ Not used as a join table
- ❌ Not used as a backend source of truth for attendance
- ❌ Not auto-synced from external systems
- ✅ Read-only for most users
- ✅ Manually maintained by designated admins

Apps Script **may read** from this sheet for:
- Display
- Email routing
- Forms dropdowns (optional)

But **no logic should break** if this sheet is temporarily out of date.

---

## Table Structure & Conventions

This sheet **does not require** strict table enforcement, but SHOULD follow SHAMROCK conventions where reasonable.

### Recommended Pattern

- Row 1: Human headers
- One row per individual
- Grouping by role type (Cadre vs Cadets) using blank rows or section headers

Unlike other tabs:
- No machine-header row is required
- Formatting may be decorative and expressive

---

## Recommended Columns

### Core Columns

| Column | Header | Description |
|------|-------|-------------|
| A | Role / Position | Official title or billet |
| B | Rank / Prefix | e.g., Col, Lt Col, Capt, C/Col |
| C | Last Name | Surname |
| D | First Name | Given name |
| E | Display Name | How they should be referenced publicly |
| F | Email | Official contact email |
| G | Office Phone | Detachment or office phone |
| H | Cell Phone | Optional |
| I | Office Location | Building / room |
| J | Notes | Optional clarifications |

---

## Example Dummy Data

| Role / Position | Rank | Last | First | Display Name | Email | Office Phone | Cell | Office | Notes |
|----------------|------|------|-------|--------------|-------|--------------|------|--------|-------|
| Detachment Commander | Col | Smith | John | Col Smith | jsmith@nd.edu | 574-631-XXXX | | ROTC HQ | |
| Operations Officer | Lt Col | Johnson | Emily | Lt Col Johnson | ejohnson@nd.edu | 574-631-XXXX | | ROTC HQ | |
| Education Officer | Capt | Lee | Michael | Capt Lee | mlee@nd.edu | 574-631-XXXX | | ROTC HQ | |
|  |  |  |  |  |  |  |  |  |  |
| Wing Commander | C/Col | Doe | Alex | C/Col Doe | adoe@nd.edu | | 574-555-1234 | | Cadet Wing |
| Vice Wing Commander | C/Lt Col | Patel | Riya | C/Lt Col Patel | rpatel@nd.edu | | | | |
| Group Commander (Blue) | C/Maj | Nguyen | Chris | C/Maj Nguyen | cnguyen@nd.edu | | | | |
| Group Commander (Gold) | C/Maj | Rivera | Sofia | C/Maj Rivera | srivera@nd.edu | | | | |
| Flight Commander (A) | C/Capt | Brown | Liam | C/Capt Brown | lbrown@nd.edu | | | | |
| Flight Commander (B) | C/Capt | Kim | Hannah | C/Capt Kim | hkim@nd.edu | | | | |

---

## Display Name Rules

- `Display Name` is the **preferred public reference**
- Used in:
  - FAQs
  - Emails
  - Forms
- Format is flexible (e.g., “Capt Lee”, “C/Maj Nguyen”)

Backend logic **must not depend** on Display Name.

---

## Editing & Governance

- This sheet is **manually maintained**
- Changes should be infrequent and deliberate
- Recommended editors:
  - Detachment staff
  - Designated SHAMROCK admins

Cadets should **not** edit this tab.

---

## Formatting Guidance (Encouraged)

Admins may use:
- Bold headers
- Section dividers (Cadre / Cadets)
- Shading by role type
- Frozen header rows
- Column grouping

Readability > uniformity.

---

## Dependencies

This sheet may be referenced by:
- FAQs (contact instructions)
- Excusals (commander selection)
- Admin menus (optional)

However:
- No other sheet should *require* this tab to function

---

## Non-Goals

The `Cadre & Leadership` tab does **not**:
- Track attendance
- Track cadet status
- Replace the Directory
- Encode permissions or authority levels

It exists to answer one question clearly:

> **“Who is responsible, and how do I contact them?”**

---

**End of `CadreLeadership.PUBLIC.md`**