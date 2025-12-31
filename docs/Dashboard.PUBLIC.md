# SHAMROCK — Dashboard (Public Specification)

**Document:** `Dashboard.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `Dashboard`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

---

## Purpose

The `Dashboard` tab is the **read-only operational overview** for SHAMROCK.

It serves as:
- The primary entry point for users
- A navigation hub to other SHAMROCK sheets
- A display surface for summaries, charts, and derived helper tables

The Dashboard **does not store source-of-truth data**.

---

## Editing Rules

- ❌ No direct data entry
- ❌ No manual row edits
- ✅ Admins may adjust layout, charts, and links
- ✅ Data is populated via formulas or Apps Script only

---

## Sections

### 1. Quick Links

Human-facing navigation links to:
- Directory
- Attendance
- Events
- Excusals
- Cadre & Leadership
- Admin Overrides
- Audit / Changelogs

This section is free-form and manually maintained.

---

### 2. Key Metrics (Summary Cells)

Examples:
- Active Cadets
- Current Training Week Attendance %
- Excusals Pending Review
- Events This Week

Values are derived from backend sheets.

---

### 3. Attendance Charts

Charts may include:
- Attendance by Flight
- Attendance by Squadron
- Attendance over Time

Charts:
- Reference hidden helper ranges
- Are read-only
- Must not be used as logic inputs

---

## Birthdays Helper Table

### Purpose

Provide a calendar-sorted birthday view for cadets:
- Sorted by month/day (ignoring year)
- Grouped into week buckets
- Formatted for announcements

### Table Structure

#### Columns

| Column | Header | Description |
|------|-------|-------------|
| A | Last Name | From Directory |
| B | First Name | From Directory |
| C | Birthday | Full date |
| D | Sorted | MM/DD key |
| E | Display | Announcement string |
| F | Group | Week bucket |

#### Example

| Last | First | Birthday | Sorted | Display | Group |
|-----|------|----------|--------|---------|-------|
| Doe | John | 1/3/2007 | 01/03 | C/Doe (1/3) | 1 |
| Cadet1 | Test | 1/10/2005 | 01/10 | C/Cadet1 (1/10) | 2 |
| Cadet2 | Test2 | 1/12/2006 | 01/12 | C/Cadet2 (1/12) | 3 |
| Cadet3 | Test3 | 1/17/2005 | 01/17 | C/Cadet3 (1/17) | 3 |

### Rules

- Directory is the source of truth
- No manual edits allowed
- Sorting and grouping are automated
- Table may be hidden or collapsed outside of birthday windows

---

## Hidden Helpers

The Dashboard may contain hidden helper ranges used for:
- Chart data
- Aggregations
- Derived metrics

Helpers must be:
- Hidden
- Protected
- Non-editable by standard users

---

## Non-Goals

The Dashboard does **not**:
- Accept data input
- Override backend records
- Serve as an audit log
- Replace any operational sheet

It exists solely to **summarize and surface** SHAMROCK data.

---

**End of `Dashboard.PUBLIC.md`**