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

---

## Project Conventions

- **No manual edits** to raw intake or system-of-record tabs
- Attendance is treated as a **transaction log**, not a wide spreadsheet
- Public excusal visibility is limited to *who / what / status*
- Excusal reasons and documents are restricted and handled separately
- Column positions should not be hardcoded; use header-based mapping helpers

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