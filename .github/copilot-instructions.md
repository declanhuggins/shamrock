
# Copilot Instructions — SHAMROCK

This repository implements **SHAMROCK** (System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping).

These instructions exist to make AI coding agents immediately productive **without redesigning the system**.

---

## 1. Big Picture Architecture

SHAMROCK uses a **strict two-tier data model**:

### Public Sheets (Views + Inputs)
- Human-facing
- Limited editability
- Never authoritative
- Defined by `*.PUBLIC.md` specs in `/docs`

### Backend Spreadsheet (Source of Truth)
- Private
- Write-only by Apps Script
- Never edited manually
- Defined by `BACKEND.SPEC.md`

All writes flow:
```
Public Sheet / Form / Menu
        ↓
   Apps Script
        ↓
 Backend Sheets
        ↓
 Rebuild Public Views
```

Public sheets **never write to each other directly**.

---

## 2. Source-of-Truth Rules (Critical)

| Data | Authority |
|----|----------|
Directory | `directory_backend` |
Events | `events_backend` |
Attendance | `attendance_backend` |
Excusals | `excusals_backend` |
Admin Actions | Apps Script menus |
Audit | `audit_log` |

If logic conflicts with a public sheet, the backend **always wins**.

---

## 3. Identity Model

- **Primary key everywhere:** `cadet_email`
- Names are cosmetic only
- Never join on names
- Never assume names are unique
- Email is immutable

Breaking this rule will corrupt history.

---

## 4. Repo Structure

```
/
├─ src/                # TypeScript source (Apps Script)
│  ├─ ingest/          # Form handlers
│  ├─ sync/            # Backend → public rebuilds
│  ├─ lib/             # Shared utilities
│  └─ main.ts          # Menus, triggers, orchestration
├─ dist/               # Compiled JS (clasp pushes ONLY this)
├─ docs/
│  ├─ *.PUBLIC.md      # Public sheet contracts
│  └─ BACKEND.SPEC.md  # Backend contract
├─ appsscript.json     # Copied into dist at build
└─ package.json
```

Do **not** push `src/` to Apps Script.

---

## 5. Apps Script Conventions (Non-Standard)

### Menus Are the Primary Admin UI
- No admin input tables
- All overrides via menus
- Menus must validate inputs
- Menus must log actions

### Tables Use Hidden Machine Headers
- Row 1 = machine keys (hidden)
- Row 2 = human headers
- Row 3+ = data
- Never reorder machine headers without code changes

---

## 6. Write Safety Rules

Every write must:
1. Acquire `LockService`
2. Batch reads/writes
3. Be idempotent
4. Emit exactly one audit log entry

No silent mutations.

---

## 7. Attendance Model

- Events exist independently of attendance
- Attendance rows created only when events are `Published`
- Cancelling an event sets attendance to `N/A`
- Never delete attendance rows

Canonical attendance codes are defined in `BACKEND.SPEC.md`.

---

## 8. Excusals Model

- Full excusal data lives in backend only
- Public excusals sheet is sanitized
- Attendance is affected only via backend logic
- Decisions must be auditable

---

## 9. Audit Log (Do Not Break This)

- Append-only
- No edits
- No deletes
- One row per mutation
- Includes actor, action, entity, old/new values, source, run_id

Audit integrity matters more than convenience.

---

## 10. Performance Expectations

- No per-row API calls
- Prefer snapshot rebuilds
- Cache lookups where safe
- Optimize for clarity over micro-performance

SHAMROCK is an administrative system, not a real-time system.

---

## 11. What NOT to Do

- Do not redesign workflows
- Do not merge backend sheets
- Do not add implicit behavior
- Do not rely on formatting or charts
- Do not weaken audit guarantees

If unsure, **follow the spec exactly**.

---

## 12. Final Rule

If implementation decisions are ambiguous:
1. Re-read `BACKEND.SPEC.md`
2. Re-read the relevant `*.PUBLIC.md`
3. Choose correctness and auditability over convenience

This system is intentionally explicit.
