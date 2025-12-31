
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  const setupMenu = ui
    .createMenu("Setup")
    .addItem("Quickstart: Backend + Forms", "shamrockQuickstart")
    .addItem("Install Backend Tabs", "shamrockInstallBackendTabs")
    .addItem("Create Cadet Intake Form", "shamrockCreateCadetForm")
    .addItem("Create Attendance Form", "shamrockCreateAttendanceForm")
    .addItem("Set Frontend Spreadsheet ID", "shamrockPromptFrontendId")
    .addItem("Open Backend Spreadsheet", "shamrockOpenBackend");

  const syncMenu = ui
    .createMenu("Sync / Rebuild")
    .addItem("Full Sync (Backend → Public)", "shamrockSyncPublicViews")
    .addSeparator()
    .addItem("Rebuild Directory", "shamrockRebuildDirectory")
    .addItem("Rebuild Attendance", "shamrockRebuildAttendance")
    .addItem("Rebuild Events", "shamrockRebuildEvents")
    .addItem("Rebuild Excusals", "shamrockRebuildExcusals")
    .addItem("Rebuild Audit", "shamrockRebuildAudit")
    .addItem("Rebuild Data Legend", "shamrockRebuildDataLegend");

  const automationMenu = ui
    .createMenu("Automation")
    .addItem("Install Daily Sync Trigger (01:00)", "shamrockInstallDailyTrigger")
    .addItem("Remove Daily Sync Triggers", "shamrockRemoveDailyTrigger");

  const utilitiesMenu = ui
    .createMenu("Utilities")
    .addItem("Import Cadets CSV (Drive)", "shamrockPromptImportCadetsCsv")
    .addItem("Simulate Cadet Intake (5)", "shamrockSimulateCadetIntake")
    .addItem("Seed Sample Events/Attendance", "shamrockSeedSampleData")
    .addItem("Seed Full Dummy Dataset", "shamrockSeedFullDummyData")
    .addItem("Run Health Check", "shamrockHealthCheck");

  ui
    .createMenu("SHAMROCK Admin")
    .addSubMenu(setupMenu)
    .addSubMenu(syncMenu)
    .addSubMenu(automationMenu)
    .addSubMenu(utilitiesMenu)
    .addToUi();
}

function logInfo(action: string, message: string, meta?: Record<string, unknown>): void {
  if (typeof Shamrock.logInfo === "function") {
    Shamrock.logInfo(action, message, meta);
    return;
  }
  try {
    Logger.log(`[SHAMROCK] ${action}: ${message}`);
  } catch (err) {
    // ignore log failures
  }
}

function logWarn(action: string, message: string, meta?: Record<string, unknown>): void {
  if (typeof Shamrock.logWarn === "function") {
    Shamrock.logWarn(action, message, meta);
    return;
  }
  try {
    Logger.log(`[SHAMROCK][WARN] ${action}: ${message}`);
  } catch (err) {
    // ignore log failures
  }
}

function logError(action: string, message: string, meta?: Record<string, unknown>): void {
  if (typeof Shamrock.logError === "function") {
    Shamrock.logError(action, message, meta);
    return;
  }
  try {
    Logger.log(`[SHAMROCK][ERROR] ${action}: ${message}`);
  } catch (err) {
    // ignore log failures
  }
}

const SHAMROCK_BACKEND_SPREADSHEET_ID = "SHAMROCK_BACKEND_SPREADSHEET_ID";
const SHAMROCK_BACKEND_ID = SHAMROCK_BACKEND_SPREADSHEET_ID;
const SHAMROCK_DIRECTORY_FORM_ID = "SHAMROCK_DIRECTORY_FORM_ID";
const SHAMROCK_ATTENDANCE_FORM_ID = "SHAMROCK_ATTENDANCE_FORM_ID";

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
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("syncAllPublicViews", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.syncAllPublicViews();
    });
    return;
  }
  logInfo("syncAllPublicViews", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.syncAllPublicViews();
  logInfo("syncAllPublicViews", "completed");
}

function shamrockQuickstart(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("quickstart", () => {
      Shamrock.ensureBackendSheets();
      logInfo("quickstart", "backend sheets ensured");
      Shamrock.rebuildDataLegend();
      logInfo("quickstart", "data legend rebuilt");
      shamrockInstallBackendTabs();
      logInfo("quickstart", "backend tabs installed");
      shamrockCreateCadetForm();
      logInfo("quickstart", "cadet form created");
      shamrockCreateAttendanceForm();
      logInfo("quickstart", "attendance form created");
      Shamrock.syncAllPublicViews();
      logInfo("quickstart", "completed backend install, forms, and sync");
    });
  } else {
    logInfo("quickstart", "begin");
    Shamrock.ensureBackendSheets();
    logInfo("quickstart", "backend sheets ensured");
    Shamrock.rebuildDataLegend();
    logInfo("quickstart", "data legend rebuilt");
    shamrockInstallBackendTabs();
    logInfo("quickstart", "backend tabs installed");
    shamrockCreateCadetForm();
    logInfo("quickstart", "cadet form created");
    shamrockCreateAttendanceForm();
    logInfo("quickstart", "attendance form created");
    Shamrock.syncAllPublicViews();
    logInfo("quickstart", "completed backend install, forms, and sync");
  }
  const ui = SpreadsheetApp.getUi();
  ui.alert("Quickstart complete", "Backend tabs, forms, and public views have been initialized. Verify the frontend ID and run a full sync if needed.", ui.ButtonSet.OK);
}

function shamrockRebuildDirectory(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildDirectory", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildDirectory();
    });
    return;
  }
  logInfo("rebuildDirectory", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildDirectory();
  logInfo("rebuildDirectory", "completed");
}

function shamrockRebuildEvents(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildEvents", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildEvents();
    });
    return;
  }
  logInfo("rebuildEvents", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildEvents();
  logInfo("rebuildEvents", "completed");
}

function shamrockRebuildAttendance(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildAttendance", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildAttendance();
    });
    return;
  }
  logInfo("rebuildAttendance", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildAttendance();
  logInfo("rebuildAttendance", "completed");
}

function shamrockRebuildExcusals(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildExcusals", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildExcusals();
    });
    return;
  }
  logInfo("rebuildExcusals", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildExcusals();
  logInfo("rebuildExcusals", "completed");
}

function shamrockRebuildAudit(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildAudit", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildAudit();
    });
    return;
  }
  logInfo("rebuildAudit", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildAudit();
  logInfo("rebuildAudit", "completed");
}

function shamrockRebuildDataLegend(): void {
  if (typeof Shamrock.withTiming === "function") {
    Shamrock.withTiming("rebuildDataLegend", () => {
      Shamrock.ensureBackendSheets();
      Shamrock.rebuildDataLegend();
    });
    return;
  }
  logInfo("rebuildDataLegend", "begin");
  Shamrock.ensureBackendSheets();
  Shamrock.rebuildDataLegend();
  logInfo("rebuildDataLegend", "completed");
}

function shamrockPromptFrontendId(): void {
  logInfo("setFrontendId", "prompting user for frontend spreadsheet id");
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Set Frontend Spreadsheet ID", "Paste the Spreadsheet ID or URL for SHAMROCK — Frontend", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const text = resp.getResponseText();
  const id = extractSpreadsheetId(text);
  if (!id) {
    ui.alert("Invalid spreadsheet ID or URL");
    logWarn("setFrontendId", "invalid input");
    return;
  }
  Shamrock.setFrontendSpreadsheetId(id);
  logInfo("setFrontendId", `saved id ${id}`);
  ui.alert("Frontend spreadsheet ID saved.");
}

function extractSpreadsheetId(idOrUrl: string): string | null {
  const match = String(idOrUrl).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function shamrockInstallDailyTrigger(): void {
  logInfo("installDailyTrigger", "scheduling daily sync at 01:00");
  removeTriggersFor("shamrockSyncPublicViews");
  ScriptApp.newTrigger("shamrockSyncPublicViews").timeBased().everyDays(1).atHour(1).create();
  logInfo("installDailyTrigger", "installed");
}

function shamrockRemoveDailyTrigger(): void {
  logInfo("removeDailyTriggers", "removing all daily sync triggers");
  removeTriggersFor("shamrockSyncPublicViews");
  logInfo("removeDailyTriggers", "removed");
}

function shamrockOpenBackend(): void {
  const backend = openBackendSpreadsheetSafe();
  const url = backend.getUrl();
  const ui = SpreadsheetApp.getUi();
  ui.alert("Backend", `Opening backend spreadsheet:\n${url}`, ui.ButtonSet.OK);
  SpreadsheetApp.getActive().toast("Opening backend spreadsheet…", "SHAMROCK");
  try {
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().getRange("A1").setNote("Backend opened: " + url);
  } catch (err) {
    // ignore toast side-effect
  }
}

function removeTriggersFor(fnName: string): void {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function shamrockPromptImportCadetsCsv(): void {
  logInfo("importCadetsCsv", "prompt start");
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Import Cadets CSV", "Paste a Drive file ID or URL for the cadet CSV.", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const id = extractSpreadsheetId(resp.getResponseText()) || resp.getResponseText().trim();
  if (!id) {
    ui.alert("No file ID detected. Please try again.");
    return;
  }
  logInfo("importCadetsCsv", `parsed id ${id}`);
  try {
    const imported = shamrockImportCadetsCsv(id);
    ui.alert(`Imported ${imported} cadets from CSV.`);
    logInfo("importCadetsCsv", `success count=${imported}`);
  } catch (err) {
    ui.alert(`Import failed: ${err}`);
    logInfo("importCadetsCsv", `failed: ${err}`);
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

function shamrockOnAttendanceFormSubmit(e: any): void {
  Shamrock.onFormSubmit(e);
}

const CADET_FORM_CONFIRMATION_MESSAGE =
  "Your response has been recorded. You can edit your information at any time by either editing the form sent to your email, or filling out the form again.";

const ATTENDANCE_FORM_CONFIRMATION_MESSAGE =
  "Thanks for submitting attendance. Remember to select yourself and verify the Training Week and Event before submitting.";

function applyCadetFormPolicies(form: GoogleAppsScript.Forms.Form): void {
  safeFormPolicy(form, "setCollectEmail", () => form.setCollectEmail(true));
  safeFormPolicy(form, "setRequireLogin", () => form.setRequireLogin(true));
  safeFormPolicy(form, "setAllowResponseEdits", () => form.setAllowResponseEdits(true));
  safeFormPolicy(form, "setShowLinkToRespondAgain", () => form.setShowLinkToRespondAgain(false));
  safeFormPolicy(form, "setConfirmationMessage", () => form.setConfirmationMessage(CADET_FORM_CONFIRMATION_MESSAGE));
}

function applyAttendanceFormPolicies(form: GoogleAppsScript.Forms.Form): void {
  safeFormPolicy(form, "setCollectEmail", () => form.setCollectEmail(true));
  safeFormPolicy(form, "setRequireLogin", () => form.setRequireLogin(true));
  safeFormPolicy(form, "setAllowResponseEdits", () => form.setAllowResponseEdits(true));
  safeFormPolicy(form, "setShowLinkToRespondAgain", () => form.setShowLinkToRespondAgain(false));
  safeFormPolicy(form, "setConfirmationMessage", () => form.setConfirmationMessage(ATTENDANCE_FORM_CONFIRMATION_MESSAGE));
}

function safeFormPolicy(form: GoogleAppsScript.Forms.Form, label: string, fn: () => void): void {
  try {
    fn();
    logInfo("formPolicy", `${label} ok`);
  } catch (err) {
    logWarn("formPolicy", `${label} skipped: ${err}`);
  }
}

function ensureCadetForm(): { form: GoogleAppsScript.Forms.Form; backend: GoogleAppsScript.Spreadsheet.Spreadsheet; responsesSheet: GoogleAppsScript.Spreadsheet.Sheet | null } {
  const backendId = getBackendIdSafe();
  const backend = openBackendSpreadsheetSafe();
  const desiredTitle = "SHAMROCK Cadet Directory Intake";

  let form: GoogleAppsScript.Forms.Form | null = openStoredCadetForm();
  logInfo("ensureCadetForm", form ? "opened stored form id" : "no stored form id found");

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
    logInfo("ensureCadetForm", "creating new form");
    form = FormApp.create(desiredTitle).setDescription("Use this form to add or update cadet directory information. Primary key: email.");
  }

  applyCadetFormPolicies(form);
  logInfo("ensureCadetForm", "applied form policies");

  const alreadyTargetingBackend = (() => {
    try {
      return form.getDestinationType() === FormApp.DestinationType.SPREADSHEET && form.getDestinationId() === (backendId || backend.getId());
    } catch (err) {
      return false;
    }
  })();

  // Ensure the form is pointing at the backend and locate the exact linked responses sheet
  if (!alreadyTargetingBackend) {
    logInfo("ensureCadetForm", "retargeting form destination to backend");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, backendId || backend.getId());
  }

  let responsesSheet = findResponseSheetForForm(backend, form) || getNewestFormResponseSheet(backend);
  if (!responsesSheet) {
    // As a fallback, search any form-linked sheet in the file
    responsesSheet = backend.getSheets().find(s => s.getFormUrl && s.getFormUrl()) || null;
  }

  responsesSheet = ensureResponseSheetName(backend, responsesSheet, "Cadet Form Responses");
  persistCadetFormId(form);
  logInfo("ensureCadetForm", "responses sheet named and form id persisted");

  return { form, backend, responsesSheet: responsesSheet || null };
}

function ensureAttendanceForm(): { form: GoogleAppsScript.Forms.Form; backend: GoogleAppsScript.Spreadsheet.Spreadsheet; responsesSheet: GoogleAppsScript.Spreadsheet.Sheet | null } {
  const backendId = getBackendIdSafe();
  const backend = openBackendSpreadsheetSafe();
  const desiredTitle = "SHAMROCK Attendance Form";

  let form: GoogleAppsScript.Forms.Form | null = openStoredAttendanceForm();
  logInfo("ensureAttendanceForm", form ? "opened stored form id" : "no stored form id found");

  const candidateSheets = [
    backend.getSheetByName("Attendance Form Responses"),
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
    logInfo("ensureAttendanceForm", "creating new form");
    form = FormApp.create(desiredTitle).setDescription("Bulk attendance capture for SHAMROCK events.");
  }

  applyAttendanceFormPolicies(form);
  logInfo("ensureAttendanceForm", "applied form policies");

  const alreadyTargetingBackend = (() => {
    try {
      return form.getDestinationType() === FormApp.DestinationType.SPREADSHEET && form.getDestinationId() === (backendId || backend.getId());
    } catch (err) {
      return false;
    }
  })();

  if (!alreadyTargetingBackend) {
    logInfo("ensureAttendanceForm", "retargeting form destination to backend");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, backendId || backend.getId());
  }

  let responsesSheet = findResponseSheetForForm(backend, form) || getNewestFormResponseSheet(backend);
  if (!responsesSheet) {
    responsesSheet = backend.getSheets().find(s => s.getFormUrl && s.getFormUrl()) || null;
  }

  responsesSheet = ensureResponseSheetName(backend, responsesSheet, "Attendance Form Responses");
  persistAttendanceFormId(form);
  logInfo("ensureAttendanceForm", "responses sheet named and form id persisted");

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
  logInfo("normalizeFormResponseSheet", `sheet=${sheet.getName()} start`);
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
  logInfo("normalizeFormResponseSheet", `sheet=${sheet.getName()} done`);
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

function openStoredAttendanceForm(): GoogleAppsScript.Forms.Form | null {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(SHAMROCK_ATTENDANCE_FORM_ID);
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

function persistAttendanceFormId(form: GoogleAppsScript.Forms.Form): void {
  try {
    const id = form.getId();
    if (id) {
      PropertiesService.getScriptProperties().setProperty(SHAMROCK_ATTENDANCE_FORM_ID, id);
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

function listActiveCadets(): any[] {
  const cadets = Shamrock.listCadets();
  return cadets.filter((c: any) => {
    const status = String((c as any).status || "").toLowerCase();
    return status === "active" || status === "";
  });
}

function formatCadetChoices(cadets: any[]): string[] {
  const sorted = [...cadets].sort((a, b) => {
    const la = String((a as any).last_name || "").toLowerCase();
    const lb = String((b as any).last_name || "").toLowerCase();
    if (la !== lb) return la < lb ? -1 : 1;
    const fa = String((a as any).first_name || "").toLowerCase();
    const fb = String((b as any).first_name || "").toLowerCase();
    return fa < fb ? -1 : fa > fb ? 1 : 0;
  });
  return sorted.map(c => {
    const last = (c as any).last_name || "";
    const first = (c as any).first_name || "";
    return `${last} ${first}`.trim();
  });
}

function groupCadetsByFlight(cadets: any[], orderedFlights: string[]): Map<string, any[]> {
  const map = new Map<string, any[]>();
  orderedFlights.forEach(f => map.set(f, []));
  map.set("Unassigned", []);
  cadets.forEach((c: any) => {
    const flight = (c as any).flight || "";
    const target = orderedFlights.find(f => f.toLowerCase() === String(flight || "").toLowerCase()) || "Unassigned";
    const list = map.get(target) || [];
    list.push(c);
    map.set(target, list);
  });
  return map;
}

function isCrossTownCadet(cadet: any): boolean {
  const dorm = String((cadet as any).dorm || "").toLowerCase();
  const uni = String((cadet as any).university || "").toLowerCase();
  if (dorm.includes("cross")) return true;
  if (uni && uni !== "notre dame") return true;
  return false;
}

function shamrockCreateCadetForm(): void {
  logInfo("createCadetForm", "begin");
  const setup = ensureCadetForm();
  const form = setup.form;
  const backend = setup.backend;

  form.setTitle("Cadet Directory Information Form");
  form.setDescription(
    "Please fill out the following form to be included in the SHAMROCK Directory. This information will be used by the Detachment for organizational, administrative, and communication purposes. Make sure all entries are accurate and up to date. You can fill out the form again or edit the form sent to your email to update any information at a later point in time.\n\nIf you need any support in filling out information in this form (such as the CIP code) or require additional options, reach out to dhuggin2@nd.edu with your inquiry."
  );
  applyCadetFormPolicies(form);

  // Refresh questions (idempotent update)
  logInfo("createCadetForm", "clearing existing items");
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
  // Keep a visible email field as a fallback, but collect verified email via Google Forms settings
  form.addTextItem().setTitle("Email").setRequired(false).setValidation(emailValidation).setHelpText("Captured automatically; only use if different.");
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
    logInfo("createCadetForm", "normalizing response sheet headers");
    normalizeFormResponseSheet(linkedSheet, getCadetFormHeaders());
  }

  removeTriggersFor("shamrockOnCadetFormSubmit");
  ScriptApp.newTrigger("shamrockOnCadetFormSubmit").forSpreadsheet(backend).onFormSubmit().create();

  const url = form.getPublishedUrl();
  updateDirectoryFormLink(url);
  logInfo("createCadetForm", "published url set and dashboard link updated");
  const ui = SpreadsheetApp.getUi();
  ui.alert("Cadet Intake Form ready", "Quick link updated on Dashboard (Directory Form).", ui.ButtonSet.OK);
}

function shamrockCreateAttendanceForm(): void {
  logInfo("createAttendanceForm", "begin");
  const setup = ensureAttendanceForm();
  const form = setup.form;
  const backend = setup.backend;

  form.setTitle("SHAMROCK Attendance Form");
  form.setDescription(
    "Bulk attendance capture for Mando, LLAB, and Secondary events. Select yourself, confirm the Training Week (TW-00), and choose the correct event type. Use the flight section for Mando/LLAB and the Secondary section for any secondary event."
  );
  applyAttendanceFormPolicies(form);

  logInfo("createAttendanceForm", "clearing existing items");
  form.getItems().forEach(item => form.deleteItem(item));

  // Page 1: metadata
  form.addTextItem().setTitle("Name").setRequired(true).setHelpText("Format: Last, First. Make sure you also check your own name below.");
  form.addTextItem().setTitle("Training Week (Format as TW-00)").setRequired(true).setHelpText("Example: TW-06 (use current week of year minus 34).");
  form
    .addMultipleChoiceItem()
    .setTitle("Event")
    .setChoiceValues(["Mando", "LLAB", "Secondary"])
    .setRequired(true)
    .setHelpText("Pick the event type you are reporting attendance for.");
  form
    .addListItem()
    .setTitle("Flight")
    .setChoiceValues(getFlightOptions())
    .setRequired(true)
    .setHelpText("Select your flight. Cross-town cadets: choose your assigned flight for LLAB; use the Cross-Town list for Mando.");

  // Page 2: Mando/LLAB flight attendance
  form.addPageBreakItem().setTitle("Flight Attendance (Mando / LLAB)").setHelpText("Check everyone present for Mando or LLAB, including yourself. Use the Cross-Town list for Mando PT only.");

  const cadets = listActiveCadets();
  logInfo("createAttendanceForm", `active cadets loaded: ${cadets.length}`);
  const flights = getFlightOptions();
  const byFlight = groupCadetsByFlight(cadets, flights);
  const crossTown = cadets.filter(isCrossTownCadet);

  flights.forEach(flight => {
    const roster = byFlight.get(flight) || [];
    if (!roster.length) return;
    form
      .addCheckboxItem()
      .setTitle(`${flight} Flight Attendance`)
      .setChoiceValues(formatCadetChoices(roster))
      .setHelpText("Select everyone present from this flight.");
  });

  if (crossTown.length) {
    form
      .addCheckboxItem()
      .setTitle("Cross-Town Cadets (Mando PT)")
      .setChoiceValues(formatCadetChoices(crossTown))
      .setHelpText("Use this only for Mando PT sessions where Cross-Town cadets form a separate element.");
  }

  // Page 3: Secondary attendance (all cadets may attend)
  form
    .addPageBreakItem()
    .setTitle("Secondary Attendance")
    .setHelpText("Secondary events may include any cadet. Select all cadets present.");

  flights.forEach(flight => {
    const roster = byFlight.get(flight) || [];
    if (!roster.length) return;
    form
      .addCheckboxItem()
      .setTitle(`Secondary – ${flight} Flight`)
      .setChoiceValues(formatCadetChoices(roster))
      .setHelpText("Select everyone from this flight who attended the secondary event.");
  });

  const responseSheet = setup.responsesSheet;
  if (responseSheet) {
    ensureResponseSheetName(backend, responseSheet, "Attendance Form Responses");
  }

  removeTriggersFor("shamrockOnAttendanceFormSubmit");
  ScriptApp.newTrigger("shamrockOnAttendanceFormSubmit").forSpreadsheet(backend).onFormSubmit().create();

  persistAttendanceFormId(form);
  logInfo("createAttendanceForm", "trigger installed and form id persisted");

  const ui = SpreadsheetApp.getUi();
  ui.alert("Attendance Form ready", "Use the linked form to collect bulk attendance. The on-submit trigger was installed.", ui.ButtonSet.OK);
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
    set("Email", sample.email);
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

function shamrockSeedSampleData(): void {
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();

    const events = [
      {
        event_id: "2026S-TW01-LLAB",
        event_name: "LLAB Kickoff",
        event_type: "LLAB",
        training_week: "TW-01",
        event_date: "2026-01-15 15:00",
        event_status: "Published",
        affects_attendance: true,
        attendance_label: "TW-01 LLAB",
        expected_group: "All Cadets",
        flight_scope: "All",
        location: "Jordan Hall",
        notes: "Seeded sample LLAB",
        created_at: Shamrock.nowIso(),
        updated_at: Shamrock.nowIso(),
      },
      {
        event_id: "2026S-TW01-MANDO",
        event_name: "Mando PT",
        event_type: "Mando",
        training_week: "TW-01",
        event_date: "2026-01-16 06:00",
        event_status: "Published",
        affects_attendance: true,
        attendance_label: "TW-01 Mando",
        expected_group: "All Cadets",
        flight_scope: "All",
        location: "Rockne",
        notes: "Seeded sample PT",
        created_at: Shamrock.nowIso(),
        updated_at: Shamrock.nowIso(),
      },
      {
        event_id: "2026S-TW02-MANDO",
        event_name: "Mando PT (Wx Cancelled)",
        event_type: "Mando",
        training_week: "TW-02",
        event_date: "2026-01-23 06:00",
        event_status: "Cancelled",
        affects_attendance: true,
        attendance_label: "TW-02 Mando",
        expected_group: "All Cadets",
        flight_scope: "All",
        location: "Rockne",
        notes: "Cancelled for weather",
        created_at: Shamrock.nowIso(),
        updated_at: Shamrock.nowIso(),
      },
    ];

    events.forEach(ev => {
      Shamrock.upsertEvent(ev as any);
      Shamrock.logAudit({
        action: "events.upsert",
        target_table: "Events Backend",
        target_key: ev.event_id,
        new_value: JSON.stringify(ev),
        source: "seed",
      });
    });

    const cadets = Shamrock.listCadets();
    if (!cadets.length) return;
    const targets = cadets.slice(0, Math.min(6, cadets.length));

    const attendanceSeeds = [
      { event_id: "2026S-TW01-LLAB", codes: ["P", "P", "P", "E", "ER", "T"] },
      { event_id: "2026S-TW01-MANDO", codes: ["P", "P", "E", "P", "MU", "P"] },
      { event_id: "2026S-TW02-MANDO", codes: ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A"] },
    ];

    targets.forEach((cadet: any, idx: number) => {
      attendanceSeeds.forEach(seed => {
        const code = seed.codes[idx] || "P";
        Shamrock.setAttendance({
          cadet_email: cadet.cadet_email,
          event_id: seed.event_id,
          attendance_code: code,
          source: "seed",
          updated_at: Shamrock.nowIso(),
        } as any);
        Shamrock.logAudit({
          action: "attendance.set_code",
          target_table: "Attendance Backend",
          target_key: `${cadet.cadet_email}|${seed.event_id}`,
          event_id: seed.event_id,
          new_value: code,
          source: "seed",
        });
      });
    });

    const excusalCadet = targets[1];
    if (excusalCadet) {
      const excusalId = `EXC-${Utilities.getUuid()}`;
      const excusal = {
        excusal_id: excusalId,
        cadet_email: excusalCadet.cadet_email,
        event_id: "2026S-TW01-LLAB",
        request_timestamp: Shamrock.nowIso(),
        reason: "Travel conflict (seed)",
        decision: "Approved",
        decision_by: "cadre@example.edu",
        decision_timestamp: Shamrock.nowIso(),
        attendance_effect: "Set E",
        source: "seed",
      } as any;
      Shamrock.appendExcusal(excusal);
      Shamrock.setAttendance({
        cadet_email: excusalCadet.cadet_email,
        event_id: excusal.event_id,
        attendance_code: "E",
        source: "seed",
        updated_at: Shamrock.nowIso(),
      } as any);
      Shamrock.logAudit({
        action: "excusals.submit",
        target_table: "Excusals Backend",
        target_key: excusalId,
        new_value: JSON.stringify(excusal),
        source: "seed",
      });
      Shamrock.logAudit({
        action: "attendance.set_code",
        target_table: "Attendance Backend",
        target_key: `${excusalCadet.cadet_email}|${excusal.event_id}`,
        event_id: excusal.event_id,
        new_value: "E",
        source: "seed",
      });
    }

    try {
      Shamrock.syncAllPublicViews();
    } catch (err) {
      // Frontend not configured; skip sync
    }
  });

  const ui = SpreadsheetApp.getUi();
  ui.alert("Sample data seeded", "Events, attendance, and one excusal were seeded. Run a public sync if needed.", ui.ButtonSet.OK);
}

function shamrockSeedFullDummyData(): void {
  Shamrock.withLock(() => {
    Shamrock.ensureBackendSheets();
    const runId = `seed-${Utilities.getUuid().slice(0, 8)}`;
    const now = Shamrock.nowIso();

    const cadets = [
      { last_name: "Smith", first_name: "Alex", as_year: "AS100", graduation_year: "2029", flight: "Alpha", squadron: "Blue", university: "Notre Dame", cadet_email: "alex.smith@example.edu", phone: "555-1001", dorm: "Alumni Hall", home_town: "South Bend", home_state: "Indiana", dob: "2007-02-11", cip_broad: "14 - Engineering", cip_code: "14.0101", afsc: "11X - Pilot", flight_path_status: "Participating 1/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Johnson", first_name: "Bailey", as_year: "AS150", graduation_year: "2029", flight: "Bravo", squadron: "Gold", university: "Notre Dame", cadet_email: "bailey.johnson@example.edu", phone: "555-1002", dorm: "Dillon Hall", home_town: "Carmel", home_state: "Indiana", dob: "2007-05-22", cip_broad: "26 - Biological and Biomedical Sciences", cip_code: "26.0202", afsc: "17D - Warfighter Communications", flight_path_status: "Participating 1/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Williams", first_name: "Casey", as_year: "AS200", graduation_year: "2028", flight: "Charlie", squadron: "Blue", university: "St. Mary's", cadet_email: "casey.williams@example.edu", phone: "555-1003", dorm: "Off-Campus", home_town: "Chicago", home_state: "Illinois", dob: "2006-08-03", cip_broad: "52 - Business, Management, Marketing", cip_code: "52.0101", afsc: "17S - Cyberspace Effects", flight_path_status: "Enrolled 2/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Brown", first_name: "Drew", as_year: "AS250", graduation_year: "2028", flight: "Delta", squadron: "Gold", university: "Holy Cross", cadet_email: "drew.brown@example.edu", phone: "555-1004", dorm: "Cross-Town", home_town: "Detroit", home_state: "Michigan", dob: "2006-12-14", cip_broad: "11 - Computer and Information Sciences", cip_code: "11.0701", afsc: "62E - Developmental Engineer", flight_path_status: "Enrolled 2/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Davis", first_name: "Emery", as_year: "AS300", graduation_year: "2027", flight: "Echo", squadron: "Blue", university: "Notre Dame", cadet_email: "emery.davis@example.edu", phone: "555-1005", dorm: "Dunne Hall", home_town: "Phoenix", home_state: "Arizona", dob: "2005-03-30", cip_broad: "14 - Engineering", cip_code: "14.0901", afsc: "12X - Combat Systems Officer", flight_path_status: "Active 3/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Miller", first_name: "Finley", as_year: "AS400", graduation_year: "2026", flight: "Foxtrot", squadron: "Gold", university: "Valparaiso", cadet_email: "finley.miller@example.edu", phone: "555-1006", dorm: "Cross-Town", home_town: "Valparaiso", home_state: "Indiana", dob: "2004-09-18", cip_broad: "45 - Social Sciences", cip_code: "45.1001", afsc: "38F - Force Support", flight_path_status: "Ready 4/4", status: "Active", photo_url: "", notes: "", created_at: now, updated_at: now },
      { last_name: "Garcia", first_name: "Hayden", as_year: "AS500", graduation_year: "2025", flight: "Alpha", squadron: "Blue", university: "Notre Dame", cadet_email: "hayden.garcia@example.edu", phone: "555-1007", dorm: "Keenan Hall", home_town: "Denver", home_state: "Colorado", dob: "2003-01-12", cip_broad: "40 - Physical Sciences", cip_code: "40.0801", afsc: "31P - Security Forces", flight_path_status: "Ready 4/4", status: "Leave", photo_url: "", notes: "PCSM waiver pending", created_at: now, updated_at: now },
      { last_name: "Lee", first_name: "Jordan", as_year: "AS600", graduation_year: "2024", flight: "Bravo", squadron: "Gold", university: "Notre Dame", cadet_email: "jordan.lee@example.edu", phone: "555-1008", dorm: "Siegfried Hall", home_town: "Seattle", home_state: "Washington", dob: "2002-11-07", cip_broad: "52 - Business, Management, Marketing", cip_code: "52.0801", afsc: "15W - Weather", flight_path_status: "Ready 4/4", status: "Alumni", photo_url: "", notes: "Commissioned May 2024", created_at: now, updated_at: now },
    ];

    cadets.forEach(record => {
      Shamrock.upsertCadet(record as any);
      Shamrock.logAudit({
        action: "directory.upsert",
        target_table: "Directory Backend",
        target_key: record.cadet_email,
        new_value: JSON.stringify(record),
        source: "seed",
        run_id: runId,
      });
    });

    const events = [
      { event_id: "2026S-TW01-LLAB", event_name: "LLAB Kickoff", event_type: "LLAB", training_week: "TW-01", event_date: "2026-01-15 15:00", event_status: "Published", affects_attendance: true, attendance_label: "TW-01 LLAB", expected_group: "All Cadets", flight_scope: "All", location: "Jordan Hall", notes: "Seeded event", created_at: now, updated_at: now },
      { event_id: "2026S-TW01-MANDO", event_name: "Mando PT", event_type: "Mando", training_week: "TW-01", event_date: "2026-01-16 06:00", event_status: "Published", affects_attendance: true, attendance_label: "TW-01 Mando", expected_group: "All Cadets", flight_scope: "All", location: "Rockne", notes: "", created_at: now, updated_at: now },
      { event_id: "2026S-TW01-SEC", event_name: "Leadership Seminar", event_type: "Secondary", training_week: "TW-01", event_date: "2026-01-17 19:00", event_status: "Draft", affects_attendance: false, attendance_label: "TW-01 Secondary", expected_group: "Optional", flight_scope: "All", location: "Hesburgh", notes: "Not tracked in attendance", created_at: now, updated_at: now },
      { event_id: "2026S-TW02-LLAB", event_name: "LLAB Mission Planning", event_type: "LLAB", training_week: "TW-02", event_date: "2026-01-22 15:00", event_status: "Published", affects_attendance: true, attendance_label: "TW-02 LLAB", expected_group: "All Cadets", flight_scope: "All", location: "Debartolo", notes: "", created_at: now, updated_at: now },
      { event_id: "2026S-TW02-MANDO", event_name: "Mando PT", event_type: "Mando", training_week: "TW-02", event_date: "2026-01-23 06:00", event_status: "Cancelled", affects_attendance: true, attendance_label: "TW-02 Mando", expected_group: "All Cadets", flight_scope: "All", location: "Rockne", notes: "Weather cancellation", created_at: now, updated_at: now },
      { event_id: "2026S-TW03-LLAB", event_name: "LLAB FTX", event_type: "LLAB", training_week: "TW-03", event_date: "2026-01-29 14:00", event_status: "Archived", affects_attendance: true, attendance_label: "TW-03 LLAB", expected_group: "All Cadets", flight_scope: "All", location: "White Field", notes: "Prior term archived sample", created_at: now, updated_at: now },
    ];

    events.forEach(ev => {
      Shamrock.upsertEvent(ev as any);
      Shamrock.logAudit({
        action: "events.upsert",
        target_table: "Events Backend",
        target_key: ev.event_id,
        new_value: JSON.stringify(ev),
        source: "seed",
        run_id: runId,
      });
    });

    const seededCadets = Shamrock.listCadets();
    const attendanceSeeds = [
      { event_id: "2026S-TW01-LLAB", codes: ["P", "P", "T", "E", "ER", "ED", "MU", "MRS"] },
      { event_id: "2026S-TW01-MANDO", codes: ["P", "P", "P", "P", "MU", "P", "E", "T"] },
      { event_id: "2026S-TW02-LLAB", codes: ["P", "T", "P", "E", "ER", "P", "E", "U"] },
      { event_id: "2026S-TW02-MANDO", codes: ["N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"] },
      { event_id: "2026S-TW03-LLAB", codes: ["P", "E", "P", "P", "MU", "P", "ES", "ED"] },
    ];

    attendanceSeeds.forEach(seed => {
      seededCadets.forEach((cadet: any, idx: number) => {
        const code = seed.codes[idx % seed.codes.length] || "P";
        Shamrock.setAttendance({
          cadet_email: cadet.cadet_email,
          event_id: seed.event_id,
          attendance_code: code,
          source: "seed",
          updated_at: Shamrock.nowIso(),
        } as any);
        Shamrock.logAudit({
          action: "attendance.set_code",
          target_table: "Attendance Backend",
          target_key: `${cadet.cadet_email}|${seed.event_id}`,
          event_id: seed.event_id,
          new_value: code,
          source: "seed",
          run_id: runId,
        });
      });
    });

    const excusals = [
      { excusal_id: "EXC-SEED-01", cadet_email: seededCadets[1]?.cadet_email || "", event_id: "2026S-TW01-LLAB", request_timestamp: now, reason: "Exam conflict", decision: "Approved", decision_by: "cadre@example.edu", decision_timestamp: now, attendance_effect: "Set E", source: "seed" },
      { excusal_id: "EXC-SEED-02", cadet_email: seededCadets[4]?.cadet_email || "", event_id: "2026S-TW01-LLAB", request_timestamp: now, reason: "Medical", decision: "Denied", decision_by: "cadre@example.edu", decision_timestamp: now, attendance_effect: "Set ED", source: "seed" },
      { excusal_id: "EXC-SEED-03", cadet_email: seededCadets[2]?.cadet_email || "", event_id: "2026S-TW02-LLAB", request_timestamp: now, reason: "Travel", decision: "Approved", decision_by: "cadre@example.edu", decision_timestamp: now, attendance_effect: "Set E", source: "seed" },
    ].filter(ex => ex.cadet_email);

    excusals.forEach(ex => {
      Shamrock.appendExcusal(ex as any);
      Shamrock.logAudit({
        action: "excusals.submit",
        target_table: "Excusals Backend",
        target_key: ex.excusal_id,
        new_value: JSON.stringify(ex),
        source: "seed",
        run_id: runId,
      });
      const normalized = String(ex.attendance_effect || "").toLowerCase();
      let code = "";
      if (normalized.includes("er")) code = "ER";
      else if (normalized.includes("set e")) code = "E";
      else if (normalized.includes("ed")) code = "ED";
      if (code) {
        Shamrock.setAttendance({
          cadet_email: ex.cadet_email,
          event_id: ex.event_id,
          attendance_code: code,
          source: "seed",
          updated_at: Shamrock.nowIso(),
        } as any);
        Shamrock.logAudit({
          action: "attendance.set_code",
          target_table: "Attendance Backend",
          target_key: `${ex.cadet_email}|${ex.event_id}`,
          event_id: ex.event_id,
          new_value: code,
          source: "seed",
          run_id: runId,
        });
      }
    });

    const adminAction = {
      action_id: `ACT-${Utilities.getUuid().slice(0, 8)}`,
      actor_email: getActorEmail(),
      action_type: "seed.sample",
      payload_json: JSON.stringify({ run_id: runId, note: "Seeded full dummy dataset" }),
      created_at: now,
      processed_at: now,
      status: "completed",
    } as any;
    Shamrock.appendAdminAction(adminAction);
    Shamrock.logAudit({
      action: "admin.action",
      target_table: "Admin Actions",
      target_key: adminAction.action_id,
      new_value: JSON.stringify(adminAction),
      source: "seed",
      run_id: runId,
    });

    try {
      Shamrock.syncAllPublicViews();
    } catch (err) {
      // Frontend not configured; skip
    }
  });

  const ui = SpreadsheetApp.getUi();
  ui.alert("Full dummy dataset seeded", "Cadets, events, attendance, excusals, and admin action logs were seeded. Run a public sync if not auto-configured.", ui.ButtonSet.OK);
}

function shamrockHealthCheck(): void {
  const ui = SpreadsheetApp.getUi();
  const lines: string[] = [];
  let backendId: string | null = null;
  let backend: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
  try {
    backendId = getBackendIdSafe();
    backend = openBackendSpreadsheetSafe();
    lines.push(`Backend: ${backend ? backend.getName() : "(active)"}`);
    if (!backendId || backendId === SHAMROCK_BACKEND_SPREADSHEET_ID) {
      lines.push("- Backend ID not set in properties (using active spreadsheet)");
    } else {
      lines.push(`- Backend ID set: ${backendId}`);
    }
  } catch (err) {
    lines.push(`Backend: not reachable (${err})`);
  }

  const requiredSheets = [
    Shamrock.BACKEND_SHEET_NAMES.cadets,
    Shamrock.BACKEND_SHEET_NAMES.events,
    Shamrock.BACKEND_SHEET_NAMES.attendance,
    Shamrock.BACKEND_SHEET_NAMES.excusals,
    Shamrock.BACKEND_SHEET_NAMES.adminActions,
    Shamrock.BACKEND_SHEET_NAMES.audit,
    Shamrock.BACKEND_SHEET_NAMES.dataLegend,
  ];

  if (backend) {
    requiredSheets.forEach(name => {
      const exists = !!backend!.getSheetByName(name);
      lines.push(`${exists ? "✅" : "⚠"} ${name}`);
    });
  }

  try {
    const frontendId = Shamrock.getFrontendSpreadsheetId();
    lines.push(`Frontend ID set: ${frontendId}`);
  } catch (err) {
    lines.push("Frontend ID not set (set via SHAMROCK Admin → Setup → Set Frontend Spreadsheet ID)");
  }

  const triggers = ScriptApp.getProjectTriggers();
  const triggerSummary = triggers.reduce<Record<string, number>>((acc, t) => {
    const handler = t.getHandlerFunction();
    acc[handler] = (acc[handler] || 0) + 1;
    return acc;
  }, {});
  const triggerLines = Object.keys(triggerSummary).map(k => `${k}: ${triggerSummary[k]}`);
  if (triggerLines.length) {
    lines.push("Triggers:");
    triggerLines.forEach(t => lines.push(`- ${t}`));
  } else {
    lines.push("No triggers installed (install daily sync and form triggers as needed)");
  }

  const message = lines.join("\n");
  Logger.log(`[SHAMROCK][health] ${message}`);
  ui.alert("SHAMROCK Health Check", message, ui.ButtonSet.OK);
}

function shamrockInstallBackendTabs(): void {
  Shamrock.ensureBackendSheets();
  const ss = openBackendSpreadsheetSafe();
  const sheets = [
    {
      name: "Directory Backend",
      machine: ["last_name", "first_name", "as_year", "graduation_year", "flight", "squadron", "university", "cadet_email", "phone", "dorm", "home_town", "home_state", "dob", "cip_broad", "cip_code", "afsc", "flight_path_status", "status", "photo_url", "notes", "created_at", "updated_at"],
      human: ["Last Name", "First Name", "AS Year", "Graduation Year", "Flight", "Squadron", "University", "Email", "Phone", "Dorm", "Home Town", "Home State", "DOB", "CIP Broad", "CIP Code", "AFSC", "Flight Path Status", "Status", "Photo Url", "Notes", "Created At", "Updated At"],
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