# SHAMROCK

**System for Headcount & Accountability of Manpower, Readiness, Oversight, and Cadet Keeping**

SHAMROCK is the authoritative cadet directory and attendance system for AFROTC Detachment 225 (the Flyin’ Irish).  
It is designed to be **public-facing and readable**, while keeping all **logic, ingestion, and accountability** driven by Apps Script automation.

This repository contains the **Apps Script codebase**, developed locally in **VS Code with TypeScript**, and deployed to Google Apps Script using **clasp**.

---

## Big Picture

SHAMROCK follows a layered model:

```
Cadets → Google Forms → Raw Intake Sheets
                         ↓
                    Apps Script
                         ↓
              System-of-Record Tables
                         ↓
               Public Presentation Views
```

- **Forms** are the only supported input method for cadets.
- **Raw intake sheets** are append-only and never edited manually.
- **System-of-record tables** are protected and written only by Apps Script.
- **Public-facing tabs** (directory, attendance, dashboards) are derived views and formatted for readability.

This separation ensures reliability, auditability, and long-term maintainability.

---

## Repository Layout

```
shamrock/
├── src/
│   ├── main.ts        # Entry points (menus, triggers)
│   ├── ingest/        # Form intake + normalization logic
│   ├── sync/          # Syncing public-facing views
│   └── lib/           # Shared helpers and config
├── dist/              # Compiled JavaScript (pushed to Apps Script)
├── appsscript.json    # Apps Script manifest
├── tsconfig.json
├── .clasp.json
├── .claspignore
└── README.md
```

Only the contents of `dist/` (plus `appsscript.json`) are pushed to Apps Script.

---

## Development Environment

- macOS
- Node.js (LTS)
- VS Code
- TypeScript
- Google Apps Script
- `@google/clasp`
- `@types/google-apps-script`

---

## Common Commands

### Login

```bash
clasp login
```

### Build TypeScript
```bash
npm run build
```

### Push to Apps Script
```bash
npm run push
```

### Live Development (recommended)
Terminal A:
```bash
npm run watch
```

Terminal B:
```bash
npm run push:watch
```

### Open Apps Script Editor
If available in your clasp version:
```bash
clasp open
```

Otherwise:
- Open the bound Google Sheet
- Extensions → Apps Script

### Seed Dummy Data (local dev/testing)
- In the bound Sheet: `SHAMROCK Admin → Utilities → Seed Full Dummy Dataset`
- This seeds cadets, events, attendance (including cancelled/archived), excusals, and an admin action, then attempts a public sync.
- Use `Seed Sample Events/Attendance` for a lighter fixture or `Simulate Cadet Intake (5)` for cadet-only seeds.

### Health Check
- In the bound Sheet: `SHAMROCK Admin → Utilities → Run Health Check`
- Verifies backend/front-end IDs, required backend sheets, and installed triggers.

---

## Project Conventions

- **No manual edits** to raw intake or system-of-record tabs
- Attendance is treated as a **transaction log**, not a wide spreadsheet
- Public excusal visibility is limited to *who / what / status*
- Excusal reasons and documents are restricted and handled separately
- Column positions should not be hardcoded; use header-based mapping helpers

## Logging and Audit

- All scripts log through `Shamrock.logInfo/Shamrock.logWarn/Shamrock.logError`, and long-running flows use `Shamrock.withTiming(action, fn)` to emit begin/end markers with durations.
- Form handlers and rebuild jobs log the target sheet/event to make it easier to trace issues in Apps Script executions.
- Every mutating path still writes a single audit row via `Shamrock.logAudit`, preserving the append-only audit trail defined in `BACKEND.SPEC.md`.
- If logging fails (e.g., Logger unavailable), core actions continue; logging is intentionally non-blocking.

---

## Deployment Notes

- This project is typically **bound to a Google Sheet** (the SHAMROCK master).
- Spreadsheet IDs and folder IDs should be stored in **Script Properties**, not hardcoded.
- Triggers (time-driven or form-driven) are configured in the Apps Script UI.

---

## Purpose

SHAMROCK exists to provide:
- A single source of truth for cadet accountability
- Transparent attendance tracking
- Clear leadership oversight
- Long-term continuity beyond any single cadet’s tenure

---

## License / Use.

The code in this repository is licensed under the MIT License. Intended for internal use for AFROTC Detachment 225.