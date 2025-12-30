# Copilot Instructions (SHAMROCK)

## Big picture
- SHAMROCK is an AFROTC Det 225 system to maintain a cadet directory + attendance + excusal workflow.
- The spreadsheet is **public-facing** for cadets, but the **system-of-record** should be script-driven and protected.
- Data flow is: **Forms → Raw intake sheets → Script normalization → System-of-record tables → Public presentation views**.

## Repo layout
- `src/main.ts` contains entry points (e.g., `onOpen`) and calls into feature modules.
- `src/ingest/*` handles reading new raw rows (from Form response sheets), validating, and writing normalized rows.
- `src/sync/*` generates/updates public-safe views in the “presentation” tabs (directory/attendance dashboards).
- `src/lib/*` holds shared helpers (sheet utilities, config access, ID helpers).

## Spreadsheet architecture assumptions
- “Raw intake” tabs are append-only and must never be edited by hand.
- “System-of-record” tabs are protected and only written by Apps Script.
- “Presentation” tabs are formatted and should be treated like a website: formulas + pivots + protected ranges.

## Conventions
- Prefer **batch processing** of new raw rows (idempotent) over fragile per-edit logic.
- Mark processed raw rows with a processed flag/timestamp so re-runs are safe.
- Avoid hardcoding column indices where possible; prefer header-based mapping helpers in `src/lib/`.
- Keep any public excusal info (who/what/status) separate from private details (reason/docs).

## Developer workflow
- Build TypeScript to `dist/` then push:
  - `npm run build` → compiles `src/` to `dist/`
  - `npm run push` → uploads `dist/` only (see `.claspignore`)
- Live dev:
  - Terminal A: `npm run watch`
  - Terminal B: `npm run push:watch`
- Open Apps Script UI: `npm run open`

## External integrations
- Google Apps Script services (SpreadsheetApp, DriveApp, PropertiesService) are the primary dependencies.
- Use Script Properties (via `PropertiesService`) for spreadsheet IDs/folder IDs if the project needs to write to multiple files.

## Examples to follow
- Add new features by creating a new module in `src/ingest/` or `src/sync/` and wiring it from `src/main.ts`.
- Keep UI hooks minimal (menus/triggers) and push all logic into modules.
