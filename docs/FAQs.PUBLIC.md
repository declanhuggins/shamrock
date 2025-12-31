# SHAMROCK — FAQs (Public Specification)

**Document:** `FAQs.PUBLIC.md`  
**Scope:** Public-facing Google Sheet tab: `FAQs`  
**System:** SHAMROCK (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping)

This document defines the **intent, structure, formatting guidance, and example content** for the public-facing **FAQs** tab.

Unlike other SHAMROCK tabs, this sheet is **almost entirely manual** and optimized for **human readability**, not automation.

---

## Purpose

The `FAQs` tab exists to answer **cadet-facing and cadre-facing questions** without requiring:
- Direct admin intervention
- Repeated emails
- Tribal knowledge

This tab should:
- Explain *what SHAMROCK is*
- Explain *what cadets are expected to do*
- Explain *what common statuses and codes mean*
- Provide *clear guidance for discrepancies and edge cases*

---

## Automation Rules

- ❌ No backend automation depends on this sheet
- ❌ No joins or lookups reference this sheet
- ❌ No Apps Script writes to this sheet
- ✅ Fully manual and editable by designated admins

This sheet is **documentation**, not data.

---

## Layout & Formatting Philosophy

This sheet **does not follow the strict table schema** used elsewhere.

Admins are encouraged to use:
- Merged cells
- Section headers
- Bulleted lists
- Emoji (sparingly)
- Visual separators
- Multiple tables if helpful

The goal is **clarity over consistency**.

---

## Recommended Structure

### Section Order (Top → Bottom)

1. What is SHAMROCK?
2. General Cadet Responsibilities
3. Attendance Basics
4. Attendance Codes & Meaning
5. Events & Calendar
6. Excusals
7. Fixing Mistakes & Discrepancies
8. Common Scenarios
9. Who to Contact

---

## Example Sheet Content

Below is **example content**, not prescriptive wording.

Admins may reword freely as long as intent is preserved.

---

### 🟢 What is SHAMROCK?

**SHAMROCK** is the official system used by the detachment to track:
- Cadet roster information
- Attendance at mandatory and optional events
- Excusal requests and decisions
- Historical participation data

It replaces multiple legacy spreadsheets and manual trackers.

If it’s about **who showed up, who didn’t, and why**, it lives here.

---

### 🧭 Cadet Responsibilities

All cadets are responsible for:
- Checking the **Attendance** tab regularly
- Submitting excusal requests **before** the event when possible
- Verifying attendance marks are correct
- Reporting discrepancies promptly

Failure to monitor SHAMROCK does **not** excuse incorrect records.

---

### 📋 Attendance Basics

- Attendance is tracked **by event**
- Events are defined centrally and published to the Attendance tab
- Each event results in a single attendance mark per cadet
- Attendance may be updated via:
  - Forms
  - Admin review
  - Approved excusals
  - Authorized overrides

---

### 🧾 Attendance Codes (Legend)

| Code | Meaning |
|-----|--------|
| **P** | Present |
| **E** | Excused |
| **ES** | Excused — Sport |
| **ER** | Excusal Request Submitted (pending) |
| **ED** | Excusal Request Denied |
| **T** | Tardy |
| **U** | Unexcused |
| **UR** | Unexcused — Report Submitted |
| **MU** | Make-Up Completed |
| **MRS** | Medical Restriction / No Physical |
| **N/A** | Event Cancelled / Not Applicable |

> **Note:** Some codes may transition over time (e.g., `ER → E` or `ER → ED`).

---

### 🗓 Events & Calendar

- Events are created and scheduled by staff
- Only events marked as **Published** affect attendance
- Draft or cancelled events will not impact records
- Event details (type, date, notes) are visible in the `Events` tab

If an event is missing, it likely:
- Has not been published yet
- Was cancelled
- Does not apply to your flight or year group

---

### 📄 Excusals

Excusals must be submitted via the **official Excusal Form**.

Key points:
- Submitting a request does **not** automatically excuse you
- Pending requests appear as `ER`
- Approved requests convert to `E` or `ES`
- Denied requests convert to `ED`

Some events may still require:
- Make-up work
- Additional documentation

---

### 🛠 Fixing Attendance Errors

If you believe your attendance is incorrect:
1. Check the **Events** tab to confirm the event applies to you
2. Verify whether an excusal is pending or denied
3. If still incorrect, contact your flight leadership or cadre

Do **not**:
- Edit attendance cells unless explicitly authorized
- Submit duplicate excusal forms for the same event

---

### 🧠 Common Scenarios

**“I submitted an excusal but it still says ER.”**  
→ Your request has not been reviewed yet.

**“I was present but marked absent.”**  
→ Contact your flight commander with details.

**“The event doesn’t apply to me.”**  
→ Verify flight, squadron, and AS year filters.

---

### 📬 Who to Contact

For questions not answered here:
- Flight Commander → First stop
- Squadron leadership → Escalation
- Cadre → Final authority

Contact information is available in the **Cadre & Leadership** tab.

---

## Maintenance Guidelines

- Update FAQs at least **once per semester**
- Remove outdated references to legacy systems
- Keep language clear, direct, and unambiguous
- Avoid policy debates — link out if needed

---

## Non-Goals

The FAQs tab does **not**:
- Define policy
- Override cadre decisions
- Serve as a historical record

It exists to **reduce confusion**, not adjudicate disputes.

---

**End of `FAQs.PUBLIC.md`**