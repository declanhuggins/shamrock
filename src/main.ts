
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  const frontendMenu = ui
    .createMenu("Frontend")
    .addItem("Full Sync (Backend → Public)", "shamrockSyncPublicViews")
    .addItem("Rebuild Directory", "shamrockRebuildDirectory")
    .addItem("Rebuild Attendance", "shamrockRebuildAttendance")
    .addItem("Rebuild Events", "shamrockRebuildEvents")
    .addItem("Rebuild Excusals", "shamrockRebuildExcusals")
    .addItem("Rebuild Audit", "shamrockRebuildAudit")
    .addItem("Rebuild Data Legend", "shamrockRebuildDataLegend")
    .addSeparator()
    .addItem("Set Frontend Spreadsheet ID", "shamrockPromptFrontendId")
    .addSeparator()
    .addItem("Install Daily Sync Trigger", "shamrockInstallDailyTrigger")
    .addItem("Remove Daily Sync Triggers", "shamrockRemoveDailyTrigger");

  const backendMenu = ui
    .createMenu("Backend")
    .addItem("Create Cadet Intake Form", "shamrockCreateCadetForm")
    .addItem("Simulate Cadet Intake (5)", "shamrockSimulateCadetIntake")
    .addItem("Import Cadets CSV (Drive)", "shamrockPromptImportCadetsCsv")
    .addItem("Install Backend Tabs", "shamrockInstallBackendTabs");

  ui
    .createMenu("Admininstrator Menu")
    .addSubMenu(frontendMenu)
    .addSeparator()
    .addSubMenu(backendMenu)
    .addToUi();
}

const SHAMROCK_BACKEND_SPREADSHEET_ID = "SHAMROCK_BACKEND_SPREADSHEET_ID";
const SHAMROCK_BACKEND_ID = SHAMROCK_BACKEND_SPREADSHEET_ID;
const SHAMROCK_DIRECTORY_FORM_ID = "SHAMROCK_DIRECTORY_FORM_ID";

function getBackendIdSafe(): string | null {
  if (typeof Shamrock.getBackendSpreadsheetIdSafe === "function") return Shamrock.getBackendSpreadsheetIdSafe();
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore
  return typeof SHAMROCK_BACKEND_ID !== "undefined" ? (SHAMROCK_BACKEND_ID as string) : null;
}

function openBackendSpreadsheetSafe(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const id = getBackendIdSafe();
  if (!id || id === "SHAMROCK_BACKEND_SPREADSHEET_ID") return SpreadsheetApp.getActive();
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    return SpreadsheetApp.getActive();
  }
}

function shamrockSyncPublicViews(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.syncAllPublicViews();
}

function shamrockRebuildDirectory(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildDirectory();
}

function shamrockRebuildEvents(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildEvents();
}

function shamrockRebuildAttendance(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildAttendance();
}

function shamrockRebuildExcusals(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildExcusals();
}

function shamrockRebuildAudit(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildAudit();
}

function shamrockRebuildDataLegend(): void {
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildDataLegend();
}

function shamrockPromptFrontendId(): void {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Set Frontend Spreadsheet ID", "Paste the Spreadsheet ID or URL for SHAMROCK — Frontend", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const text = resp.getResponseText();
  const id = extractSpreadsheetId(text);
  if (!id) {
    ui.alert("Invalid spreadsheet ID or URL");
    return;
  }
  Shamrock.setFrontendSpreadsheetId(id);
  ui.alert("Frontend spreadsheet ID saved.");
}

function extractSpreadsheetId(idOrUrl: string): string | null {
  const match = String(idOrUrl).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function shamrockInstallDailyTrigger(): void {
  removeTriggersFor("shamrockSyncPublicViews");
  ScriptApp.newTrigger("shamrockSyncPublicViews").timeBased().everyDays(1).atHour(1).create();
}

function shamrockRemoveDailyTrigger(): void {
  removeTriggersFor("shamrockSyncPublicViews");
}

function removeTriggersFor(fnName: string): void {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function shamrockPromptImportCadetsCsv(): void {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Import Cadets CSV", "Paste a Drive file ID or URL for the cadet CSV.", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const id = extractSpreadsheetId(resp.getResponseText()) || resp.getResponseText().trim();
  if (!id) {
    ui.alert("No file ID detected. Please try again.");
    return;
  }
  try {
    const imported = shamrockImportCadetsCsv(id);
    ui.alert(`Imported ${imported} cadets from CSV.`);
  } catch (err) {
    ui.alert(`Import failed: ${err}`);
  }
}

function shamrockImportCadetsCsv(fileId: string): number {
  return Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const setup = ensureCadetForm();
    const responsesSheet = setup.responsesSheet;
    if (!responsesSheet) throw new Error("Cadet Form Responses sheet not found");
    normalizeFormResponseSheet(responsesSheet, getCadetFormHeaders());

    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    return importCadetsFromCsvContent(content, responsesSheet);
  });
}

function importCadetsFromCsvContent(content: string, responsesSheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const rows = Utilities.parseCsv(content);
  if (!rows.length) return 0;

  const machineHeaders = rows[0].map(h => Shamrock.normalizeHeader(h));
  const headerMap = Shamrock.buildHeaderMap(machineHeaders);
  const hasHumanHeader = rows.length > 1 && rows[1].some(cell => String(cell || "").toLowerCase().includes("first"));
  const dataStart = hasHumanHeader ? 2 : 1;

  const targetHeaders = getCadetFormMachineHeaders();
  const targetLen = targetHeaders.length;
  const appendRows: any[][] = [];
  let imported = 0;

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    const emailIdx = headerMap["cadet_email"];
    const email = emailIdx != null ? String(row[emailIdx] || "").trim().toLowerCase() : "";
    if (!email) continue;

    const responseRow = targetHeaders.map(field => {
      const idx = headerMap[field];
      const raw = idx != null ? row[idx] : "";
      if (raw && typeof raw === "object" && Object.prototype.toString.call(raw) === "[object Date]") {
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        return (raw as Date).toISOString();
      }
      return String(raw ?? "");
    });

    const tsIdx = targetHeaders.indexOf("timestamp");
    if (tsIdx !== -1 && !responseRow[tsIdx]) {
      const createdIdx = headerMap["created_at"];
      const createdVal = createdIdx != null ? row[createdIdx] || "" : "";
      responseRow[tsIdx] = createdVal || new Date().toISOString();
    }

    appendRows.push(responseRow);

    const record: Record<string, any> = {};
    (Shamrock.CADET_FIELDS as string[]).forEach((field: string) => {
      const idx = headerMap[field];
      record[field] = idx != null ? row[idx] || "" : "";
    });
    record.cadet_email = email;
    record.created_at = record.created_at || Shamrock.nowIso();
    record.updated_at = Shamrock.nowIso();

    Shamrock.upsertCadet(record as any);
    imported++;
  }

  if (appendRows.length) {
    const startRow = Math.max(responsesSheet.getLastRow() + 1, 3);
    const neededRows = startRow + appendRows.length - 1;
    if (responsesSheet.getMaxRows() < neededRows) {
      responsesSheet.insertRowsAfter(responsesSheet.getMaxRows(), neededRows - responsesSheet.getMaxRows());
    }
    const targetRange = responsesSheet.getRange(startRow, 1, appendRows.length, targetLen);
    // Clear any existing validations on target rows to avoid conflicts with dropdown/range rules
    targetRange.clearDataValidations();
    try {
      targetRange.setValues(appendRows);
    } catch (err) {
      targetRange.clearDataValidations();
      targetRange.setValues(appendRows);
    }
  }

  return imported;
}

function toNamedRange(name: string): string {
  let sanitized = name.replace(/[^A-Za-z0-9_]/g, "_");
  if (!/^[A-Za-z]/.test(sanitized)) {
    sanitized = `NR_${sanitized}`;
  }
  if (!sanitized.length) {
    sanitized = "NR_default";
  }
  return sanitized;
}

function removeNamedRangeIfExists(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, name: string): void {
  ss.getNamedRanges()
    .filter(nr => nr.getName() === name)
    .forEach(nr => nr.remove());
}

function getDataLegendOptionRange(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, key: string): GoogleAppsScript.Spreadsheet.Range | null {
  const sheet = ss.getSheetByName(Shamrock.BACKEND_SHEET_NAMES.dataLegend);
  if (!sheet) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] as string[];
  const colIndex = headers.indexOf(key);
  const lastRow = sheet.getLastRow();
  if (colIndex === -1 || lastRow < 3) return null;
  return sheet.getRange(3, colIndex + 1, lastRow - 2, 1);
}

function applyDirectoryBackendValidations(sheetName: string, sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  if (sheetName !== "Directory Backend") return;

  // Ensure at least one data row exists for validations
  if (sheet.getMaxRows() < 3) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 3 - sheet.getMaxRows());
  }
  const dataRowCount = Math.max(sheet.getMaxRows() - 2, 1);

  const ss = sheet.getParent();
  const validators: Array<{ col: number; range?: GoogleAppsScript.Spreadsheet.Range | null; values?: string[]; email?: boolean }> = [
    { col: 3, range: getDataLegendOptionRange(ss, "as_year_options"), values: getAsYearOptions() },
    { col: 5, range: getDataLegendOptionRange(ss, "flight_options"), values: getFlightOptions() },
    { col: 6, range: getDataLegendOptionRange(ss, "squadron_options"), values: getSquadronOptions() },
    { col: 7, range: getDataLegendOptionRange(ss, "university_options"), values: getUniversityOptions() },
    { col: 8, range: getDataLegendOptionRange(ss, "dorm_options"), values: getDormOptions() },
    { col: 9, email: true },
    { col: 12, range: getDataLegendOptionRange(ss, "home_state_options"), values: getHomeStateOptions() },
    { col: 14, range: getDataLegendOptionRange(ss, "afsc_options"), values: getAfscOptions() },
    { col: 15, range: getDataLegendOptionRange(ss, "cip_broad_options"), values: getCipBroadOptions() },
    { col: 16, values: getCipCodeOptions() },
    { col: 17, range: getDataLegendOptionRange(ss, "flight_path_status_options"), values: getFlightPathStatusOptions() },
    { col: 18, range: getDataLegendOptionRange(ss, "status_options"), values: getStatusOptions() },
  ];

  validators.forEach(v => {
    const range = sheet.getRange(3, v.col, dataRowCount, 1);
    if (v.range) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(v.range, true).setAllowInvalid(false).build();
      range.setDataValidation(rule);
    } else if (v.values) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(v.values, true).setAllowInvalid(false).build();
      range.setDataValidation(rule);
    } else if (v.email) {
      const rule = SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build();
      range.setDataValidation(rule);
      applyPeopleChips(sheet, v.col, dataRowCount);
    }
  });
}

function applyPeopleChips(sheet: GoogleAppsScript.Spreadsheet.Sheet, col: number, dataRows: number): void {
  try {
    if (typeof Sheets === "undefined" || !Sheets.Spreadsheets) return;
    const sheetId = sheet.getSheetId();
    // Use a loose type to avoid schema drift across Apps Script versions
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [
      {
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: 2,
            endRowIndex: 2 + dataRows,
            startColumnIndex: col - 1,
            endColumnIndex: col,
          },
          cell: { userEnteredDataType: { type: "PERSON" } } as unknown as GoogleAppsScript.Sheets.Schema.CellData,
          fields: "userEnteredDataType",
        },
      },
    ];
    const backendId = getBackendIdSafe();
    if (!backendId || backendId === "SHAMROCK_BACKEND_SPREADSHEET_ID") return;
    Sheets.Spreadsheets.batchUpdate({ requests }, backendId);
  } catch (err) {
    // If PEOPLE chips are unsupported, fall back silently
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function shamrockOnCadetFormSubmit(e: any): void {
  Shamrock.onFormSubmit(e);
}

const CADET_FORM_CONFIRMATION_MESSAGE =
  "Your response has been recorded. You can edit your information at any time by either editing the form sent to your email, or filling out the form again.";

function applyCadetFormPolicies(form: GoogleAppsScript.Forms.Form): void {
  form
    .setCollectEmail(true)
    .setRequireLogin(true)
    .setAllowResponseEdits(true)
    .setShowLinkToRespondAgain(false)
    .setConfirmationMessage(CADET_FORM_CONFIRMATION_MESSAGE);
}

function ensureCadetForm(): { form: GoogleAppsScript.Forms.Form; backend: GoogleAppsScript.Spreadsheet.Spreadsheet; responsesSheet: GoogleAppsScript.Spreadsheet.Sheet | null } {
  const backendId = getBackendIdSafe();
  const backend = openBackendSpreadsheetSafe();
  const desiredTitle = "SHAMROCK Cadet Directory Intake";

  let form: GoogleAppsScript.Forms.Form | null = openStoredCadetForm();

  // Start with likely candidates so we preserve an existing form when possible
  const candidateSheets = [
    backend.getSheetByName("Cadet Form Responses"),
    backend.getSheetByName("Form Responses 1"),
    backend.getSheetByName("Form Responses"),
  ].filter((s): s is GoogleAppsScript.Spreadsheet.Sheet => Boolean(s));

  let formUrl: string | null = null;
  if (!form) {
    for (const sheet of candidateSheets) {
      try {
        if (sheet.getFormUrl && sheet.getFormUrl()) {
          formUrl = sheet.getFormUrl();
          break;
        }
      } catch (err) {
        formUrl = null;
      }
    }
    if (formUrl) {
      form = FormApp.openByUrl(formUrl);
    }
  }

  if (!form) {
    form = FormApp.create(desiredTitle).setDescription("Use this form to add or update cadet directory information. Primary key: university email.");
  }

  applyCadetFormPolicies(form);

  const alreadyTargetingBackend = (() => {
    try {
      return form.getDestinationType() === FormApp.DestinationType.SPREADSHEET && form.getDestinationId() === (backendId || backend.getId());
    } catch (err) {
      return false;
    }
  })();

  // Ensure the form is pointing at the backend and locate the exact linked responses sheet
  if (!alreadyTargetingBackend) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, backendId || backend.getId());
  }

  let responsesSheet = findResponseSheetForForm(backend, form) || getNewestFormResponseSheet(backend);
  if (!responsesSheet) {
    // As a fallback, search any form-linked sheet in the file
    responsesSheet = backend.getSheets().find(s => s.getFormUrl && s.getFormUrl()) || null;
  }

  responsesSheet = ensureResponseSheetName(backend, responsesSheet, "Cadet Form Responses");
  persistCadetFormId(form);

  return { form, backend, responsesSheet: responsesSheet || null };
}

function updateDirectoryFormLink(url: string): void {
  try {
    const frontendId = Shamrock.getFrontendSpreadsheetId();
    const frontend = SpreadsheetApp.openById(frontendId);
    const sheet = frontend.getSheetByName(Shamrock.PUBLIC_SHEET_NAMES.dashboard);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (String(values[r][c]).trim().toLowerCase() === "directory form") {
          sheet.getRange(r + 1, c + 2).setFormula(`=HYPERLINK("${url}","Open Form")`);
          return;
        }
      }
    }
  } catch (err) {
    // If frontend not configured or sheet missing, quietly skip
  }
}

function normalizeFormResponseSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet, headers?: string[]): void {
  const expectedHuman = headers && headers.length ? headers : null;
  const machineHeaders = getCadetFormMachineHeaders();
  const targetCols = machineHeaders.length;

  // Ensure at least two header rows exist
  if (sheet.getMaxRows() < 2) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 2 - sheet.getMaxRows());
  }

  // Drop any surplus columns entirely, then ensure we have enough columns
  const maxBefore = sheet.getMaxColumns();
  if (maxBefore > targetCols) {
    sheet.deleteColumns(targetCols + 1, maxBefore - targetCols);
  }
  if (sheet.getMaxColumns() < targetCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCols - sheet.getMaxColumns());
  }

  const maxCols = sheet.getMaxColumns();

  // If the machine header row is not already present, insert a row above the human headers
  const firstRow = sheet.getRange(1, 1, 1, maxCols).getValues()[0];
  const machineMatches = machineHeaders.length === maxCols && machineHeaders.every((val, idx) => String(firstRow[idx] ?? "") === val);
  if (!machineMatches) {
    sheet.insertRowBefore(1);
  }

  // Machine headers row 1 (pad to full width)
  sheet.getRange(1, 1, 1, targetCols).setValues([machineHeaders]);

  // Human headers row 2 (form titles, padded)
  if (expectedHuman) {
    const humanRow = expectedHuman.slice(0, targetCols).concat(Array(Math.max(0, targetCols - expectedHuman.length)).fill(""));
    sheet.getRange(2, 1, 1, targetCols).setValues([humanRow]);
  }

  // Delete any rows beyond the two header rows (form will grow new rows for responses)
  let maxRows = sheet.getMaxRows();
  // Keep at least one data row to avoid deleting all non-frozen rows
  if (maxRows > 3) {
    sheet.deleteRows(4, maxRows - 3);
    maxRows = sheet.getMaxRows();
  } else if (maxRows < 3) {
    sheet.insertRowsAfter(maxRows, 3 - maxRows);
    maxRows = sheet.getMaxRows();
  }

  // Clear any lingering validations on the data rows to avoid conflicts when importing rows
  if (maxRows >= 3) {
    const dataRows = maxRows - 2;
    sheet.getRange(3, 1, dataRows, targetCols).clearDataValidations();
  }

  // Name the human header row to match the sheet (sanitized for named range rules), and hide machine headers
  const ss = sheet.getParent();
  const rangeName = toNamedRange(sheet.getName());
  removeNamedRangeIfExists(ss, rangeName);
  ss.setNamedRange(rangeName, sheet.getRange(2, 1, 1, targetCols));

  // Keep machine headers out of view and freeze both header rows for clarity
  sheet.hideRows(1);
  sheet.setFrozenRows(2);
}

function getCadetFormHeaders(): string[] {
  return Shamrock.CADET_FORM_HEADERS;
}

function getCadetFormMachineHeaders(): string[] {
  return Shamrock.CADET_FORM_MACHINE_HEADERS;
}

function getAsYearOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.as_year_options;
}

function getUniversityOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.university_options;
}

function getSquadronOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.squadron_options;
}

function getFlightOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.flight_options;
}

function getHomeStateOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.home_state_options;
}

function getFlightPathStatusOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.flight_path_status_options;
}

function getStatusOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.status_options;
}

function getDormOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.dorm_options;
}

function getCipBroadOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.cip_broad_options;
}

function getCipCodeOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.cip_code_options;
}

function getAfscOptions(): string[] {
  return Shamrock.DATA_LEGEND_OPTIONS.afsc_options;
}

function findResponseSheetForForm(backend: GoogleAppsScript.Spreadsheet.Spreadsheet, form: GoogleAppsScript.Forms.Form): GoogleAppsScript.Spreadsheet.Sheet | null {
  const matchUrls: string[] = [];
  try {
    matchUrls.push(form.getEditUrl());
  } catch (err) {
    // ignore
  }
  try {
    matchUrls.push(form.getPublishedUrl());
  } catch (err) {
    // ignore
  }
  if (!matchUrls.length) return null;

  for (const sheet of backend.getSheets()) {
    try {
      const sheetFormUrl = sheet.getFormUrl && sheet.getFormUrl();
      if (sheetFormUrl && matchUrls.indexOf(sheetFormUrl) !== -1) {
        return sheet;
      }
    } catch (err) {
      // ignore and continue
    }
  }
  return null;
}

function ensureResponseSheetName(
  backend: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  desiredName: string,
): GoogleAppsScript.Spreadsheet.Sheet | null {
  if (!sheet) return null;
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const targetId = sheet.getSheetId();
  const isFormResponsesName = (name: string) => /^Form Responses( \d+)?$/i.test(name);

  backend.getSheets().forEach(candidate => {
    const name = candidate.getName();
    const isTarget = candidate.getSheetId() === targetId;
    const isConflict = name === desiredName || isFormResponsesName(name);
    if (!isTarget && isConflict) {
      candidate.setName(`${name} (old ${stamp}-${candidate.getSheetId()})`);
    }
  });

  sheet.setName(desiredName);
  return sheet;
}

function openStoredCadetForm(): GoogleAppsScript.Forms.Form | null {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(SHAMROCK_DIRECTORY_FORM_ID);
    if (!id) return null;
    return FormApp.openById(id);
  } catch (err) {
    return null;
  }
}

function persistCadetFormId(form: GoogleAppsScript.Forms.Form): void {
  try {
    const id = form.getId();
    if (id) {
      PropertiesService.getScriptProperties().setProperty(SHAMROCK_DIRECTORY_FORM_ID, id);
    }
  } catch (err) {
    // ignore persistence failures
  }
}

function getNewestFormResponseSheet(backend: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Spreadsheet.Sheet | null {
  const formSheets = backend.getSheets().filter(s => {
    try {
      return Boolean(s.getFormUrl && s.getFormUrl());
    } catch (err) {
      return false;
    }
  });
  if (!formSheets.length) return null;
  // Newly created form response sheets are appended at the end
  return formSheets[formSheets.length - 1];
}

function shamrockCreateCadetForm(): void {
  const setup = ensureCadetForm();
  const form = setup.form;
  const backend = setup.backend;

  form.setTitle("Cadet Directory Information Form");
  form.setDescription(
    "Please fill out the following form to be included in the SHAMROCK Directory. This information will be used by the Detachment for organizational, administrative, and communication purposes. Make sure all entries are accurate and up to date. You can fill out the form again or edit the form sent to your email to update any information at a later point in time.\n\nIf you need any support in filling out information in this form (such as the CIP code) or require additional options, reach out to dhuggin2@nd.edu with your inquiry."
  );
  applyCadetFormPolicies(form);

  // Refresh questions (idempotent update)
  form.getItems().forEach(item => form.deleteItem(item));
  const emailValidation = FormApp.createTextValidation().requireTextIsEmail().build();
  const phoneValidation = FormApp.createTextValidation().requireTextMatchesPattern("^\\d{10}$").setHelpText("Enter 10 digits, numbers only").build();
  const gradYearValidation = FormApp.createTextValidation().requireTextMatchesPattern("^20\\d{2}$").setHelpText("Use 4 digits, e.g., 2027").build();
  const cipCodeValidation = FormApp.createTextValidation().requireTextMatchesPattern("^\\d{2}\\.\\d{4}$").setHelpText("Format: 12.3456").build();

  form.addTextItem().setTitle("Last Name").setRequired(true);
  form.addTextItem().setTitle("First Name").setRequired(true);
  form
    .addListItem()
    .setTitle("AS Year")
    .setChoiceValues(getAsYearOptions())
    .setRequired(true)
    .setHelpText(
      "If you are in your first of 4 years in AFROTC, you are an AS100. If you are in your first of 3 planned years in AFROTC, you are an AS250. AS150 = joined in spring; AS500 = PSP non-select returning; AS700/AS800 = extended cadets (non-scholarship/scholarship); AS900 = completed AFROTC awaiting commission."
    );
  form.addTextItem().setTitle("Graduation Year").setValidation(gradYearValidation).setHelpText("e.g., 2027");
  form.addListItem().setTitle("Flight").setChoiceValues(getFlightOptions()).setRequired(true);
  form.addListItem().setTitle("Squadron").setChoiceValues(getSquadronOptions()).setRequired(true);
  form.addListItem().setTitle("University").setChoiceValues(getUniversityOptions()).setRequired(true);
  form
    .addListItem()
    .setTitle("Dorm")
    .setChoiceValues(getDormOptions())
    .setRequired(false)
    .setHelpText("Answer \"Off-Campus\" or \"Cross-Town\" if not applicable.");
  form.addTextItem().setTitle("University Email").setRequired(true).setValidation(emailValidation);
  form.addTextItem().setTitle("Phone").setValidation(phoneValidation).setHelpText("Type your 10 digit phone number without any other characters.");
  form.addTextItem().setTitle("Home Town");
  form.addListItem().setTitle("Home State").setChoiceValues(getHomeStateOptions());
  form.addDateItem().setTitle("DOB");
  form.addListItem().setTitle("AFSC").setChoiceValues(getAfscOptions());
  form.addListItem().setTitle("CIP Broad").setChoiceValues(getCipBroadOptions()).setHelpText("Find CIP codes at nces.ed.gov/ipeds/cipcode/browse.aspx?y=56");
  form.addTextItem().setTitle("CIP Code").setValidation(cipCodeValidation).setHelpText("Short answer, format NN.NNNN e.g., 11.0701");
  form
    .addListItem()
    .setTitle("Flight Path Status")
    .setChoiceValues(getFlightPathStatusOptions())
    .setRequired(true)
    .setHelpText("Will be stored with 1/4, 2/4, etc.");
  form.addListItem().setTitle("Status").setChoiceValues(getStatusOptions()).setRequired(true).setHelpText("Defaults to Active if blank");
  form.addTextItem().setTitle("Photo Link").setHelpText("Optional: link to a photo (URL)");
  form.addParagraphTextItem().setTitle("Notes").setHelpText("Optional: any additional notes");

  const linkedSheet = setup.responsesSheet;
  if (linkedSheet) {
    normalizeFormResponseSheet(linkedSheet, getCadetFormHeaders());
  }

  removeTriggersFor("shamrockOnCadetFormSubmit");
  ScriptApp.newTrigger("shamrockOnCadetFormSubmit").forSpreadsheet(backend).onFormSubmit().create();

  const url = form.getPublishedUrl();
  updateDirectoryFormLink(url);
  const ui = SpreadsheetApp.getUi();
  ui.alert("Cadet Intake Form ready", "Quick link updated on Dashboard (Directory Form).", ui.ButtonSet.OK);
}

function shamrockSimulateCadetIntake(): void {
  const { form } = ensureCadetForm();
  const items = form.getItems();
  const itemByTitle: Record<string, GoogleAppsScript.Forms.Item> = {};
  items.forEach(it => {
    const textItem = (it as any).asTextItem ? (it as any).asTextItem() : null;
    if (textItem) {
      const title = textItem.getTitle();
      if (title) itemByTitle[title] = it;
    }
  });

  const samples = [
    { email: "cadet1@example.edu", last: "Smith", first: "Alex", as_year: "AS200", flight: "Alpha", squadron: "1", university: "Notre Dame", dorm: "Dorm A", phone: "555-1001", dob: "2003-01-01", cip_broad: "Engineering", cip_code: "14.0101", afsc: "Pilot", fps: "On Track", status: "Active" },
    { email: "cadet2@example.edu", last: "Johnson", first: "Blake", as_year: "AS300", flight: "Bravo", squadron: "2", university: "Notre Dame", dorm: "Dorm B", phone: "555-1002", dob: "2002-02-02", cip_broad: "Science", cip_code: "26.0101", afsc: "Intel", fps: "On Track", status: "Active" },
    { email: "cadet3@example.edu", last: "Williams", first: "Casey", as_year: "AS100", flight: "Charlie", squadron: "3", university: "Notre Dame", dorm: "Dorm C", phone: "555-1003", dob: "2004-03-03", cip_broad: "Business", cip_code: "52.0101", afsc: "Finance", fps: "Exploring", status: "Active" },
    { email: "cadet4@example.edu", last: "Brown", first: "Drew", as_year: "AS400", flight: "Delta", squadron: "4", university: "Notre Dame", dorm: "Dorm D", phone: "555-1004", dob: "2001-04-04", cip_broad: "Humanities", cip_code: "23.0101", afsc: "PA", fps: "On Track", status: "Active" },
    { email: "cadet5@example.edu", last: "Davis", first: "Emery", as_year: "AS200", flight: "Echo", squadron: "5", university: "Notre Dame", dorm: "Dorm E", phone: "555-1005", dob: "2003-05-05", cip_broad: "Science", cip_code: "26.0201", afsc: "Cyber", fps: "On Track", status: "Active" },
  ];

  samples.forEach(sample => {
    const response = form.createResponse();
    const set = (title: string, value: string) => {
      const item = itemByTitle[title];
      if (item && (item as any).getType && (item as any).getType() === FormApp.ItemType.TEXT && (item as any).asTextItem) {
        response.withItemResponse((item as any).asTextItem().createResponse(value));
      }
    };
    set("University Email", sample.email);
    set("Last Name", sample.last);
    set("First Name", sample.first);
    set("AS Year", sample.as_year);
    set("Flight", sample.flight);
    set("Squadron", sample.squadron);
    set("University", sample.university);
    set("Dorm", sample.dorm);
    set("Phone", sample.phone);
    set("DOB", sample.dob);
    set("CIP Broad", sample.cip_broad);
    set("CIP Code", sample.cip_code);
    set("AFSC", sample.afsc);
    set("Flight Path Status", sample.fps);
    set("Status", sample.status);
    response.submit();
  });
}

function shamrockInstallBackendTabs(): void {
  Shamrock.ensureBackendSheets();
  const ss = openBackendSpreadsheetSafe();
  const sheets = [
    {
      name: "Directory Backend",
      machine: ["last_name", "first_name", "as_year", "graduation_year", "flight", "squadron", "university", "dorm", "cadet_email", "phone", "home_town", "home_state", "dob", "afsc", "cip_broad", "cip_code", "flight_path_status", "status", "photo_url", "notes", "created_at", "updated_at"],
      human: ["Last Name", "First Name", "AS Year", "Graduation Year", "Flight", "Squadron", "University", "Dorm", "University Email", "Phone", "Home Town", "Home State", "DOB", "AFSC", "CIP Broad", "CIP Code", "Flight Path Status", "Status", "Photo Url", "Notes", "Created At", "Updated At"],
    },
    {
      name: "Events Backend",
      machine: ["event_id", "event_name", "event_type", "training_week", "event_date", "event_status", "affects_attendance", "attendance_label", "expected_group", "flight_scope", "location", "notes", "created_at", "updated_at"],
      human: ["Event ID", "Event Name", "Event Type", "Training Week", "Event Date", "Event Status", "Affects Attendance", "Attendance Label", "Expected Group", "Flight Scope", "Location", "Notes", "Created At", "Updated At"],
    },
    {
      name: "Attendance Backend",
      machine: ["cadet_email", "event_id", "attendance_code", "source", "updated_at"],
      human: ["Cadet Email", "Event ID", "Attendance Code", "Source", "Updated At"],
    },
    {
      name: "Excusals Backend",
      machine: ["excusal_id", "cadet_email", "event_id", "request_timestamp", "reason", "decision", "decision_by", "decision_timestamp", "attendance_effect", "source"],
      human: ["Excusal ID", "Cadet Email", "Event ID", "Request Timestamp", "Reason", "Decision", "Decision By", "Decision Timestamp", "Attendance Effect", "Source"],
    },
    {
      name: "Admin Actions",
      machine: ["action_id", "actor_email", "action_type", "payload_json", "created_at", "processed_at", "status"],
      human: ["Action ID", "Actor Email", "Action Type", "Payload (JSON)", "Created At", "Processed At", "Status"],
    },
    {
      name: "Audit Log",
      machine: ["audit_id", "timestamp", "actor_email", "actor_role", "action", "target_sheet", "target_table", "target_key", "target_range", "event_id", "request_id", "old_value", "new_value", "result", "reason", "notes", "source", "script_version", "run_id"],
      human: ["Audit ID", "Timestamp", "Actor Email", "Actor Role", "Action", "Target Sheet", "Target Table", "Target Key", "Target Range", "Event ID", "Request ID", "Old Value", "New Value", "Result", "Reason", "Notes", "Source", "Script Version", "Run ID"],
    },
  ];

  sheets.forEach(({ name, machine, human }) => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    // Machine headers (Row 1)
    sheet.getRange(1, 1, 1, machine.length).setValues([machine]);
    // Human headers (Row 2)
    sheet.getRange(2, 1, 1, human.length).setValues([human]);
    // Trim extra columns beyond machine headers
    const maxCols = sheet.getMaxColumns();
    if (maxCols > machine.length) {
      sheet.deleteColumns(machine.length + 1, maxCols - machine.length);
    }
    // Trim extra rows beyond header rows
    const maxRows = sheet.getMaxRows();
    if (maxRows > 2) {
      sheet.deleteRows(3, maxRows - 2);
    }
    // Name row 2 as a range (sanitized) matching the sheet name and hide machine headers (row 1)
    const rangeName = toNamedRange(name);
    removeNamedRangeIfExists(ss, rangeName);
    ss.setNamedRange(rangeName, sheet.getRange(2, 1, 1, machine.length));
    sheet.hideRows(1);

    // Apply dropdown chips / people chips on data rows starting at row 3
    applyDirectoryBackendValidations(name, sheet);
  });
}

// Expose form submit handler
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function onFormSubmit(e: any): void {
  Shamrock.onFormSubmit(e);
}