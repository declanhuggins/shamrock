# SHAMROCK Backend Forms & Sheet Contract

Authoritative reference for how Google Forms (or other intake surfaces) write into backend sheets. This supplements BACKEND.SPEC.md; if conflicts arise, BACKEND.SPEC.md wins.

## Core Rules
- All writes acquire LockService and emit one audit_log row.
- Row 1 = machine headers (do not edit or reorder). Row 2 = human headers (optional in backend). Row 3+ = data.
- Primary keys are immutable: cadet_email for cadets/attendance, event_id for events, excusal_id for excusals.
- Apps Script is the only writer; do not edit backend sheets manually.

## Sheet Contracts & Form Flows

### directory_backend (Directory intake)
- **Primary key:** cadet_email (university email).
- **Fields (machine header order):** cadet_email, last_name, first_name, as_year, flight, squadron, university, dorm, phone, dob, cip_broad, cip_code, afsc, flight_path_status, status, created_at, updated_at.
- **Form mapping:**
  - Required: cadet_email (from "University Email" / "Email" field).
  - Optional fields above map 1:1 by name; missing values become empty strings.
  - status defaults to "Active" if omitted.
  - created_at/updated_at auto-populated on submission.
- **Behavior:**
  - Upsert by cadet_email; a new submission will replace stored values (blanking fields that are not present).
  - Audit entry action: `directory.upsert`, target_table `directory_backend`, target_key = cadet_email.

### events_backend (Event intake)
- **Primary key:** event_id.
- **Fields:** event_id, event_name, event_type, training_week, event_date, event_status, affects_attendance, attendance_label, expected_group, flight_scope, location, notes, created_at, updated_at.
- **Form mapping:**
  - event_id required; if missing, generated as `TERM-TW-TYPE` or `EVT-<uuid>`.
  - event_status defaults to "Published"; affects_attendance defaults to TRUE; flight_scope defaults to "All"; attendance_label defaults to event_name or event_id.
  - created_at/updated_at auto-populated.
- **Behavior:**
  - Upsert by event_id; latest submission overwrites prior values when fields are absent.
  - Attendance rows are created later via rebuild/menus, not by the form handler itself.
  - Audit entry action: `events.upsert`, target_table `events_backend`, target_key = event_id.

### attendance_backend (Indirect writes)
- **Composite key:** cadet_email + event_id.
- **Fields:** cadet_email, event_id, attendance_code, source, updated_at.
- **Write paths:**
  - Not written by a form directly; set via menu actions or excusal effects.
  - Cancelled events set attendance_code to "N/A" during public sync rebuild.
- **Audit:** action `attendance.set_code`, target_key `${cadet_email}|${event_id}`.

### excusals_backend (Excusal intake)
- **Primary key:** excusal_id (generated `EXC-<uuid>` on submit).
- **Fields:** excusal_id, cadet_email, event_id, request_timestamp, reason, decision, decision_by, decision_timestamp, attendance_effect, source.
- **Form mapping:**
  - Required: cadet_email, event_id.
  - attendance_effect defaults to "Set ER" if omitted.
  - decision/decision_by optional; decision_timestamp set only when decision is present.
  - request_timestamp auto-populated.
- **Behavior:**
  - Always appends (no updates in place).
  - Immediately applies attendance_effect to attendance_backend (ER/E/ED codes only) and logs audit for both excusal submission and attendance change.
  - Audit entry actions: `excusals.submit` (target_key = excusal_id) and optional `attendance.set_code`.

### admin_actions (Menu queue)
- **Fields:** action_id, actor_email, action_type, payload_json, created_at, processed_at, status.
- **Usage:** populated by menu-driven admin tools; not written by forms. Actions are append-only commands.
- **Audit:** each admin action produces an audit_log entry referencing the action_id.

### audit_log (Append-only)
- **Fields (machine header order):** audit_id, timestamp, actor_email, actor_role, action, target_sheet, target_table, target_key, target_range, event_id, request_id, old_value, new_value, result, reason, notes, source, script_version, run_id.
- **Behavior:** one row per mutation; never edit or delete.

## Expected Form Sheets
- Each Google Form lands responses in its own tab (e.g., "Cadet Form Responses", "Event Form Responses", "Excusal Form Responses"). The handler dispatches by sheet name containing `cadet`, `event`, or `excusal` (case-insensitive).
- Do not rename machine headers in backend sheets; only adjust human headers if needed.
- If a response omits required keys (cadet_email for directory, event_id or deduced ID for events, cadet_email/event_id for excusals), the submission throws and writes nothing.

## Operational Notes
- After form ingestion, run `SHAMROCK → Sync All Public Views` to refresh public sheets.
- Backend sheets must stay hidden from casual editors; treat them as logs/ledgers, not UI.
- For onboarding, create forms that expose only the fields above and validate email formats before submission.
