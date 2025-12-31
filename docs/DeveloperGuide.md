# SHAMROCK Developer Guide

This guide complements the PUBLIC specs and BACKEND contracts. It focuses on developer workflows, safety rules, and how to exercise the system locally.

## Setup
- Install dependencies: `npm install`
- Build once: `npm run build -- --pretty false`
- Set script properties in Apps Script or via menu:
  - Backend ID: SHAMROCK_BACKEND_ID or script property `SHAMROCK_BACKEND_SPREADSHEET_ID`
  - Frontend ID: menu `SHAMROCK Admin → Setup → Set Frontend Spreadsheet ID`
- Quickstart menu: `SHAMROCK Admin → Setup → Quickstart: Backend + Forms`

## Safety Checklist
1. All writes acquire `LockService` (enforced by code paths).
2. One audit row per mutation (`Shamrock.logAudit`).
3. Backend sheets are source-of-truth; public tabs are projections only.
4. Do not join on names; email (`cadet_email`) is the primary key.
5. No manual edits to backend sheets.

## Forms & Ingestion
- Form submissions are handled by `Shamrock.onFormSubmit` (see `src/ingest/forms.ts`).
- Cadet form maps `Email`/`Email Address`/`University Email` → `cadet_email` (lowercased).
- Event form prefers provided `event_id` or derives `{term}-{tw}-{type}`.
- Excusal form requires email + event ID and applies attendance effects (ER/E/ED) immediately.
- Attendance form resolves events by training week + type when IDs are absent and records `P`.

## Seeding & Fixtures
- Light fixtures: `SHAMROCK Admin → Utilities → Seed Sample Events/Attendance`
- Full dataset: `SHAMROCK Admin → Utilities → Seed Full Dummy Dataset`
  - Seeds cadets, events (including cancelled/archived and non-attendance), attendance codes, excusals, and an admin action; attempts a public sync.
- Cadet-only: `Simulate Cadet Intake (5)` fills the cadet form and triggers ingestion.

## Health Check
- Run `SHAMROCK Admin → Utilities → Run Health Check`
- Reports:
  - Backend ID presence and reachable sheet
  - Existence of required backend tabs
  - Frontend ID presence
  - Installed project triggers (daily sync, form triggers)

## Sync & Rebuilds
- Full rebuild: `SHAMROCK Admin → Sync / Rebuild → Full Sync (Backend → Public)`
- Targeted rebuilds exist for Directory, Attendance, Events, Excusals, Audit, and Data Legend.
- Cancelled events write `N/A` across attendance when rebuilt.

## Testing Notes
- Prefer running `npm run build -- --pretty false` before `npm run push`.
- The repo uses header-driven writes; avoid adding/removing columns without updating constants and specs.
- If frontend ID is missing, rebuild calls are safe no-ops for public sheets.

## Troubleshooting
- Missing cadets in attendance: check Directory backend and rerun `Rebuild Attendance`.
- Excusal not reflected: ensure event is published/affects attendance and rerun `Rebuild Attendance`/`Rebuild Excusals`.
- Trigger missing: reinstall via menus (daily sync or form creation routines).
