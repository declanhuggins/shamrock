# Public Documentation

Operational, shareable notes for the SHAMROCK Sheets/Forms system. Keep this file free of secrets (no raw sheet/form IDs, emails, or personal data). Use it to record how each feature works for day-to-day users and operators.

This file is the public-facing feature catalog. System-wide invariants and architectural rules live in internal documentation.

## How to add a feature entry
1) Copy the template in `../templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`.
2) Paste it as a new section in this file.
3) Fill in all placeholders, replacing any sensitive values with friendly names or instructions on where to find them.
4) Include validation steps so non-developers can confirm behavior after deployments.

Recommended workflow for AI agents:
- Confirm invariants in `docs/system/SYSTEM_SPEC.md`.
- Follow the checklist in `docs/ai/FEATURE_CHANGE_CHECKLIST.md`.

## Feature catalog
Maintain a simple index of features here so operators can find the right section quickly.

Index columns:
- Feature name
- Surfaces (frontend tabs, backend tabs, forms)
- Entry points (menus, triggers)
- Status
- Section link

No features documented yet.

---

## Setup / Provisioning
- **Owner/POC**: Engineering
- **Status**: active
- **Last updated**: 2025-12-31

### Overview
Runs an idempotent "ensure-exists" setup that creates/ensures the frontend workbook, backend workbook, required tabs, and the Attendance/Excusal forms. Safe to re-run; avoids duplicates.

### User entry points
- Custom menu in frontend sheets: SHAMROCK → "Run setup (ensure-exists)".
- Script editor (for admins): run global function `setup`.

### Data touched
- Workbooks: SHAMROCK Frontend, SHAMROCK Backend.
- Tabs ensured (frontend): FAQs, Dashboard, Leadership, Directory, Attendance, Events, Excusals, Data Legend.
- Tabs ensured (backend): Directory Backend, Leadership Backend, Events Backend, Excusals Backend, Attendance Backend, Audit Backend, Data Legend.
- Forms ensured:
	- SHAMROCK Attendance Form
	- SHAMROCK Excusal Form
	- SHAMROCK Directory Form (collects responder email; allows response editing)

### Workflow (happy path)
1) Operator runs setup via menu or `setup` function.
2) Script ensures (creates if missing) the two workbooks and stores their IDs in Script Properties.
3) Script ensures required tabs exist and backfills row 1 (machine headers) and row 2 (display headers) if empty.
4) Script ensures the Attendance and Excusal forms exist and collect verified responder emails.
5) Operator sees a completion alert summarizing counts.

What setup auto-runs
- Applies frontend formatting/validations (Directory/Leadership/Attendance/Data Legend/FAQs) and creates Sheets “tables” for those tabs.
- Syncs Data Legend from canonical arrays to frontend; syncs Directory from backend; rebuilds the Attendance matrix.
- Normalizes form response sheets, trims Attendance response columns, reapplies Attendance Backend formatting.
- Installs onOpen/onEdit triggers for menus and backend sync; sets up form submit triggers.

### Error handling and safeguards
- Idempotent: rerun setup to repair missing resources; it will not intentionally duplicate tabs/forms.
- If a stored ID is invalid, setup recreates the resource and updates Script Properties.
- Avoid manual renames of tabs to preserve matching; if renamed, rerun setup to recreate missing tabs.
- Logging: setup writes INFO/WARN messages to the execution logs (Apps Script console/Logger).

### Deployment / configuration
- Requires Apps Script authorization to create sheets/forms and manage properties.
- No secrets are stored; resource IDs are saved in Script Properties (environment-specific).

### Validation checklist
- After running setup, confirm the SHAMROCK menu appears.
- Open the frontend workbook and verify the listed tabs exist with two header rows.
- Open the backend workbook and verify its tabs exist with two header rows.
- Open both forms and confirm email collection is enabled and login is required.

### Known limits / open questions
- Apps Script cannot create or modify Google Sheets “Formatted tables” (typed columns). If you want to use them, you must create/maintain them manually in the Sheets UI (see checklist below).
- Attendance/Excusal form questions are placeholders; real questions will be added later.
- The built-in Forms “email a copy of my responses” setting may not be controllable via Apps Script; if needed, a submission-trigger email receipt will be implemented.

### Manual formatted-table setup (Sheets UI)
Apps Script can’t build or update Google’s new “Formatted table” objects. After running `setup` and letting the formatter run once, do this manually in Sheets if you want typed columns/colored dropdowns that stick:

1) **Freeze & headers**: Scripts freeze rows 1–2 and hide row 1 (machine headers). Keep row 1 hidden; don’t reorder columns.
2) **Directory**: Select from row 2 down (include display headers). Insert → Formatted table. Column order: Last, First, Year, Class, Flight, Squadron, University, Email, Phone, Dorm, Home Town, Home State, DOB, CIP Broad, CIP Code, Desired/Assigned AFSC, Flight Path Status, Photo, Notes.
3) **Leadership**: Select from row 2 down, insert a formatted table. Order: Last, First, Rank, Role, Email, Office Phone, Cell Phone, Office Location.
4) **Attendance**: Select from row 2 down (include events). Overall in column F, LLAB in column G. Insert formatted table. Script still applies conditional formatting for attendance codes; banding may be overwritten if `applyAll` reruns.
5) **Excusals / Data Legend / Dashboard / FAQs**: Optional formatted tables from row 2 down; keep row 1 hidden.
6) **Dropdown colors**: Set background colors on source ranges in Data Legend, then create dropdown “from a range” to inherit chips. Colors are UI-only; scripts can’t set them. Use Paste special → Data validation to reuse.
7) **Protections**: Header rows, Directory name columns, and Attendance A–G are locked; event columns are editable by Leadership emails. Table creation still works because row 2+ remain editable.
8) **Preserve visuals**: If you want to keep the table look, set script property `DISABLE_FRONTEND_FORMATTING=true` (menu: SHAMROCK → Toggle Frontend Formatting), create tables/colors, then run SHAMROCK → Reapply Frontend Protections.

### Fresh install / startup steps
Follow this order to stand up a brand-new environment:

1) Clone the repo

```bash
git clone https://github.com/declanhuggins/shamrock.git
```

2) Install dependencies

```bash
cd shamrock
npm install
```

3) Authenticate `clasp` to your Google account

```bash
clasp login
```

4) Create a new Apps Script project (standalone)

```bash
npm run create
```

5) Open the Apps Script project in the browser

```bash
Created new script: https://script.google.com...
└─ appsscript.json
Cloned one file..
```

6) Push the local build output to Apps Script

```bash
npm run push
```

7) Provision everything (first run)
- In the Apps Script editor, run the function in `index.js` called `setup`.
- Approve all scopes when prompted (Sheets, Forms, Drive, Gmail). Click `Review permissions`, select your account, clikc `Advanced`, and then `Go to Shamrock (unsafe)`. Select `Select all` and then `Continue`.
- Run `setup` one more time after auth so it can finish cleanly.

8) Add your email to admins (menu access)
- In the Apps Script editor: Project Settings → Script properties → Add property `SHAMROCK_MENU_ALLOWED_EMAILS` with your email (comma-separated list for multiple admins). Save.
- This gate controls who sees the SHAMROCK menu in the spreadsheets.

9) Confirm the Sheets UI entry point
- Open the generated frontend spreadsheet.
- Use SHAMROCK → “Run setup (ensure-exists)” and confirm it completes.

10) Apply required form settings (manual in Forms UI)
- Directory Form: Settings → Responses → “Send responders a copy of their response” = Always; “Allow response editing” = On.
- Attendance Form: “Send responders a copy” = Off; “Allow response editing” = Off.
- Excusal Form: “Send responders a copy” = Off; “Allow response editing” = Off.
- Ensure “Collect email addresses” remains On for all forms (setup sets this).

### Migrating to a new script or forms
If you need to move to a fresh Apps Script project or new Forms (e.g., to fix broken destinations):

1) Export data from the old environment
- In the existing frontend spreadsheet: SHAMROCK → “Export category (backend)”.
- Export each category you care about (at minimum: `directory`, `events`, `excusals`, `data_legend`).
- Optional: also export Directory as CSV using the canonical cadet CSV format.

2) Create a new Apps Script project
- If you’re doing this on a new machine, authenticate first:

```bash
clasp login
```

- Create the new project:

```bash
clasp create
```

- Push code:

```bash
npm run push
```

- Open the project:

```bash
clasp open
```

3) Re-run provisioning in the new project
- In the Apps Script editor, run `setup`.
- Approve scopes, then run `setup` one more time.

4) Re-import data into the new backend
- In the new frontend spreadsheet: SHAMROCK → “Import category (backend)” and import each JSON export.
- If you exported Directory as CSV, use SHAMROCK → “Import cadet CSV -> Directory Backend”.

5) Validate + cutover
- Verify each Form’s response destination is the new backend workbook.
- Verify the backend has these response sheet names:
	- Attendance Form Responses
	- Excusal Form Responses
	- Directory Form Responses
- Swap any shared links/bookmarks to point to the new forms/workbooks.
- Re-apply the form settings noted in step 9 (Form Settings → Responses) after migration.
