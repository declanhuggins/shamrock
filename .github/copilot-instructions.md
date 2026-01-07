# Copilot Instructions for SHAMROCK

- **Stack**: Google Apps Script (V8) in TypeScript, built locally (`npm run build`) and deployed with `clasp` (`npm run push`). See `appsscript.json` and `package.json`.
- **Surfaces**: Google Sheets frontends (locked, data via forms/logic) plus Google Forms with verified responder emails. System creates sheets/forms via APIs; add custom menus for user/admin actions.

## Architecture & layout
- Frontend workbook: user-facing sheets; backend workbook: source-of-truth mirrors (`Directory/Attendance/Events/Excusals/Audit/Data Legend`) with extra helper columns as needed. Backend syncs forward; avoid hardcoding column positions—use machine-friendly headers in row 1 with tables starting row 2.
- `src/` conventions: `forms/` (form handlers), `sheets/` (tab helpers, column maps), `services/` (business logic and sync), `triggers/` (entry points), `utils/`, `config/` (IDs/flags, no secrets), `types.ts` (shared contracts). Keep modules small and orchestrate in `services/`.
- Docs: `docs/AI_CONTRIBUTION_GUIDE.md` (workflow), `docs/public/README.md` (feature catalog), `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md` (structure).

## Frontend workbook requirements
- **FAQs**: simple two-column info for end users.
- **Dashboard**: quick links (other sheets/forms/GitHub), key metrics, charts, rotating upcoming birthdays (from Directory DOB), rotating cadets-out-this-week (from Attendance/Events). Use formatting, smart chips, borders.
- **Cadre & Leadership**: table columns role/position, rank, last/first, email, office phone, cell, office; minimal contact view.
- **Directory**: table sorted Z-A by AS year then A-Z last name. Columns: last name, first name, AS year, class year (YYYY), flight, squadron (hidden), university, email, phone (display +5 (555) 555-5555), dorm, home town, home state, DOB (MM/DD/YYYY), CIP broad area, CIP code (hidden), desired/assigned AFSC, flight path status, photo link, notes; allow extra hidden helper columns. Place borders between university|email, dorm|home town, DOB|CIP broad, flight path status|photo link; squadron and CIP code hidden. Normalize names/emails, trim whitespace; validate phone format. Use dropdowns from Data Legend ranges: AS years, flights, squadrons, universities, dorms, home states, CIP broad areas, CIP codes (six digits), AFSC options, flight path statuses. Accept unique name cases (e.g., “ben Yosef”).
- **Attendance**: cadet list synced from Directory (last/first/AS/flight/squadron) plus event columns with dropdown codes: P, E, ES, ER, ED, T, U, UR, MU, MRS, N/A, or blank (pending). Credit: P/E/ES/MU/MRS; No credit: U/ED; Pending: ER; Excluded: N/A. Include LLAB attendance % and overall attendance % columns.
- **Events**: event metadata to drive attendance and dashboard (e.g., Event ID, Term, Training Week, Event Type, Display Name, Attendance Column Label, Expected Group, Flight Scope, Status, Start/End datetime, Location, Notes, Created At/By).
- **Excusals**: public-facing log (Request ID, Event, Email, Last/First, Flight, Squadron, Status, Decision, Decided By/At, Attendance Effect, Submitted, Last Updated, Notes).
- **Audit/Changelog**: track changes (e.g., Audit ID, Timestamp, Actor Email, Role, Action, Target Sheet/Table/Key/Range, Event ID, Request ID, Old/New, Result, Reason, Notes, Source, Version, Run ID).
- **Data Legend**: host validation ranges (AS years, flights, squadrons, universities, dorms, states, CIP lists, AFSCs, flight path statuses, attendance codes).

## Backend workbook requirements
- Mirror frontend sheets; backend is authoritative. Directory/Events/Excusals managed here and propagated to frontend. Attendance backend is an append-only log of submissions; rebuild frontend attendance by replaying log + excusal decisions.

## Forms
- **Attendance form**: multi-page; capture responder email/name; select event (TW-XX categories). For mando, choose flight or crosstown (Trine/Valparaiso) then show cadet checkboxes grouped by AS year (only where cadets exist). For secondaries, show all cadets sorted by last name and grouped by AS year (no flight selection). Allow recording multiple cadets per submission.
- **Excusal form**: capture email/last/first, select event (dropdown), reason, PDF upload (MFR). Backend enriches flight/squadron from Directory. Decision dropdown (Approved/Denied/blank) on backend drives attendance updates and notifies cadet’s squadron/flight commanders (from Cadre & Leadership).

## Guardrails & workflows
- Keep sheets locked; edits flow through forms/logic; use menus for admin actions (e.g., regenerate attendance, sync directory, install triggers). Apply conditional formatting, borders, and smart chips where useful.
- Build: `npm run build`; deploy: `npm run push`; pull remote: `npm run pull`; clean: `npm run clean`.
- Default advanced service: Sheets API v4. Configure IDs/flags via `config/`, not hardcoded. Document new features in `docs/public/README.md` using the template and keep README in sync with scripts/structure.

