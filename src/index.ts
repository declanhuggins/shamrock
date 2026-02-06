// Entry points for SHAMROCK Apps Script.

function getAllowedMenuUsers(): string[] {
  try {
    const prop = Config.scriptProperties().getProperty('SHAMROCK_MENU_ALLOWED_EMAILS') || '';
    return prop
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter(Boolean);
  } catch (err) {
    Log.warn(`Unable to read SHAMROCK_MENU_ALLOWED_EMAILS property: ${err}`);
    return [];
  }
}

function addShamrockMenu() {
  const email = ((): string => {
    try {
      const active = Session.getActiveUser().getEmail();
      if (active) return active;
    } catch (err) {
      Log.warn(`Unable to read active user email for menu gate: ${err}`);
    }
    return '';
  })();

  const allowed = getAllowedMenuUsers();
  const emailLower = (email || '').toLowerCase();

  if (allowed.length > 0 && !allowed.includes(emailLower)) {
    Log.warn(`Menu suppressed for user=${email || 'unknown'}; not in allowed list`);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const dangerMenu = ui
    .createMenu('DANGEROURS')
    .addSubMenu(
      ui
        .createMenu('Setup & Automations')
        .addItem('Run setup (ensure-exists)', 'setup')
        .addItem('Pause automations (defer sync)', 'pauseAutomations')
        .addItem('Resume automations (process pending)', 'resumeAutomations')
        .addItem('Reinstall all triggers', 'reinstallAllTriggers')
    )
    .addSubMenu(
      ui
        .createMenu('Sync & Refresh')
        .addItem('Sync Directory (Backend -> Frontend)', 'syncDirectoryBackendToFrontend')
        .addItem('Sync Leadership (Backend -> Frontend)', 'syncLeadershipBackendToFrontend')
        .addItem('Sync Data Legend (Backend -> Frontend)', 'syncDataLegendBackendToFrontend')
        .addItem('Sync ALL mapped (Backend -> Frontend)', 'syncAllBackendToFrontend')
        .addSeparator()
        .addItem('Refresh Events + Attendance artifacts', 'refreshEventsArtifacts')
        .addItem('Refresh Data Legend (Backend -> Frontend)', 'refreshDataLegendAndFrontend')
        .addItem('Rebuild Dashboard', 'rebuildDashboard')
        .addItem('Rebuild Attendance Matrix (backend log -> frontend matrix)', 'rebuildAttendanceMatrix')
        .addItem('Rebuild Attendance Form (events -> form choices)', 'rebuildAttendanceForm')
        .addItem('Refresh Excusals Form (events -> form choices)', 'refreshExcusalsForm')
        .addItem('Setup Excusals Management Spreadsheet', 'setupExcusalsManagementSpreadsheet')
        .addItem('Share Excusals Management Spreadsheet', 'shareExcusalsManagementSpreadsheet')
        .addItem('Reinitialize Excusals Management Sheets', 'reinitializeExcusalsManagementSheets')
        .addItem('Process Excusals Form Backlog', 'processExcusalsFormBacklog')
        .addItem('Process Attendance Form Backlog', 'processAttendanceFormBacklog')
        .addItem('Prune Attendance Response Duplicates', 'pruneAttendanceResponseColumns')
        .addItem('Prune Excusals Response Duplicates', 'pruneExcusalsResponseColumns')
        .addSeparator()
        .addItem('Reorder Frontend Sheets', 'reorderFrontendSheets')
        .addItem('Reorder Backend Sheets', 'reorderBackendSheets')
    )
    .addSubMenu(
      ui
        .createMenu('Formatting & Protections')
        .addItem('Apply Frontend Formatting', 'applyFrontendFormatting')
        .addItem('Toggle Frontend Formatting (on/off)', 'toggleFrontendFormatting')
        .addItem('Toggle Column Width Formatting (on/off)', 'toggleFrontendColumnWidths')
        .addItem('Reapply Frontend Protections', 'reapplyFrontendProtections')
    )
    .addSubMenu(
      ui
        .createMenu('Imports/Exports (Backend)')
        .addItem('Export Cadets CSV (Directory Backend)', 'exportCadetsCsv')
        .addItem('Import Cadets CSV (Directory Backend)', 'importCadetsCsv')
        .addItem('Export Leadership CSV (Leadership Backend)', 'exportLeadershipCsv')
        .addItem('Import Leadership CSV (Leadership Backend)', 'importLeadershipCsv')
        .addItem('Export Events CSV (Events Backend)', 'exportEventsCsv')
        .addItem('Import Events CSV (Events Backend)', 'importEventsCsv')
        .addItem('Export Attendance CSV (Attendance Backend)', 'exportAttendanceCsv')
        .addItem('Import Attendance CSV (Attendance Backend)', 'importAttendanceCsv')
    );

  const safeMenu = ui
    .createMenu('SAFE FUNCTIONS')
    .addItem('Add Leadership Entry', 'addLeadershipEntry')
    .addItem('Fix Attendance Headers', 'fixAttendanceHeaders')
    .addItem('Fill Attendance Event (bulk)', 'fillAttendanceEventPrompt')
    .addItem('Replay Latest Directory Form', 'replayLatestDirectoryFormResponse')
    .addItem('Debug Attendance Response Sheet', 'debugAttendanceResponseSheet');

  ui
    .createMenu('SHAMROCK')
    .addSubMenu(dangerMenu)
    .addSubMenu(safeMenu)
    .addItem('Show menu help / data flow', 'showMenuHelp')
    .addToUi();
}

function onOpen() {
  try {
    addShamrockMenu();
  } catch (err) {
    Log.warn(`onOpen failed to add menu: ${err}`);
  }
}

function setup() {
  const summary = SetupService.runSetup();
  const message = [
    'Setup completed.',
    `Spreadsheets: ${summary.spreadsheets.length}`,
    `Sheets ensured: ${summary.sheets.length}`,
    `Forms: ${summary.forms.length}`,
  ].join('\n');

  // Show an alert only if a container-bound UI is available; otherwise log.
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(message);
  } catch (err) {
    Log.warn(`No UI context for alert; logging summary instead. Error: ${err}`);
    Log.info(message);
  }
}

function exportEventsCsv() {
  AdminService.exportEventsCsv();
}

function importEventsCsv() {
  AdminService.importEventsCsv();
}

function exportAttendanceCsv() {
  AdminService.exportAttendanceCsv();
}

function importAttendanceCsv() {
  AdminService.importAttendanceCsv();
}

function exportLeadershipCsv() {
  AdminService.exportLeadershipCsv();
}

function importLeadershipCsv() {
  AdminService.importLeadershipCsv();
}

function exportCadetsCsv() {
  AdminService.exportCadetsCsv();
}

function importCadetsCsv() {
  AdminService.importCadetsCsv();
}

function syncDirectoryBackendToFrontend() {
  SetupService.syncDirectoryBackendToFrontend();
}

function syncLeadershipBackendToFrontend() {
  SetupService.syncLeadershipBackendToFrontend();
}

function syncDataLegendBackendToFrontend() {
  SetupService.syncDataLegendBackendToFrontend();
}

function syncAllBackendToFrontend() {
  SetupService.syncAllBackendToFrontend();
}

function refreshDataLegendAndFrontend() {
  SetupService.refreshDataLegendAndFrontend();
}

function refreshEventsArtifacts() {
  SetupService.refreshEventsArtifacts();
}

function rebuildDashboard() {
  SetupService.rebuildDashboard();
}

function rebuildAttendanceMatrix() {
  SetupService.rebuildAttendanceMatrix();
}

function sendWeeklyMandoExcusedSummary() {
  AttendanceService.sendWeeklyMandoExcusedSummary();
}

function sendWeeklyLlabExcusedSummary() {
  AttendanceService.sendWeeklyLlabExcusedSummary();
}

function sendWeeklyUnexcusedSummary() {
  AttendanceService.fillUnexcusedAndNotify();
}

function rebuildAttendanceForm() {
  SetupService.rebuildAttendanceForm();
}

function reorderFrontendSheets() {
  SetupService.reorderFrontendSheets();
}

function reorderBackendSheets() {
  SetupService.reorderBackendSheets();
}

function applyFrontendFormatting() {
  SetupService.applyFrontendFormatting();
}

function pauseAutomations() {
  SetupService.pauseAutomations();
}

function resumeAutomations() {
  SetupService.resumeAutomations();
}

function toggleFrontendFormatting() {
  SetupService.toggleFrontendFormatting();
}

function toggleFrontendColumnWidths() {
  SetupService.toggleFrontendColumnWidths();
}

function reapplyFrontendProtections() {
  SetupService.reapplyFrontendProtections();
}

function archiveCoreSheets() {
  SetupService.archiveCoreSheets();
}

function restoreCoreSheetsFromArchive() {
  SetupService.restoreCoreSheetsFromArchive();
}

function pruneExcusalsResponseColumns() {
  SetupService.pruneExcusalsResponseColumns();
}

function refreshExcusalsForm() {
  SetupService.refreshExcusalsForm();
}

function processExcusalsFormBacklog() {
  SetupService.processExcusalsFormBacklog();
}

function processAttendanceFormBacklog() {
  SetupService.processAttendanceFormBacklog();
  SpreadsheetApp.getUi().alert('Processed attendance form backlog. Check logs for details.');
}

function debugAttendanceResponseSheet() {
  const headers = SetupService.debugAttendanceResponseSheet();
  SpreadsheetApp.getUi().alert(`Found ${headers.length} columns in Attendance Response Sheet. Check logs for details.`);
}

function setupExcusalsManagementSpreadsheet() {
  try {
    const managementId = ExcusalsService.ensureManagementSpreadsheet();
    ExcusalsService.shareAndProtectManagementSpreadsheet();
    const url = `https://docs.google.com/spreadsheets/d/${managementId}`;
    SpreadsheetApp.getUi().alert(`Excusals management spreadsheet ready and shared:\n${url}`);
  } catch (err) {
    SpreadsheetApp.getUi().alert(`Failed to set up management spreadsheet: ${err}`);
  }
}

function shareExcusalsManagementSpreadsheet() {
  try {
    ExcusalsService.shareAndProtectManagementSpreadsheet();
    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');
    const url = managementId ? `https://docs.google.com/spreadsheets/d/${managementId}` : 'N/A';
    SpreadsheetApp.getUi().alert(`Excusals management spreadsheet shared with commanders and protected:\n${url}`);
  } catch (err) {
    SpreadsheetApp.getUi().alert(`Failed to share management spreadsheet: ${err}`);
  }
}

function reinitializeExcusalsManagementSheets() {
  try {
    ExcusalsService.ensureManagementSpreadsheet();
    ExcusalsService.shareAndProtectManagementSpreadsheet();
    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');
    const url = managementId ? `https://docs.google.com/spreadsheets/d/${managementId}` : 'N/A';
    SpreadsheetApp.getUi().alert(`Excusals management sheets reinitialized and protected:\n${url}`);
  } catch (err) {
    SpreadsheetApp.getUi().alert(`Failed to reinitialize management sheets: ${err}`);
  }
}

function debugExcusalsResponseColumnsVerbose() {
  SetupService.debugExcusalsResponseColumnsVerbose();
}

function reinstallAllTriggers() {
  SetupService.reinstallAllTriggers();
}

function addLeadershipEntry() {
  // Prompt for basic leadership fields and append to Leadership Backend.
  try {
    const ui = SpreadsheetApp.getUi();
    const ask = (label: string, required = false): string | null => {
      const res = ui.prompt(label, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return null;
      const value = String(res.getResponseText() || '').trim();
      if (required && !value) return ask(label, required); // re-prompt if required and empty
      return value;
    };

    const lastName = ask('Last Name', true);
    if (lastName === null) return;
    const firstName = ask('First Name', true);
    if (firstName === null) return;
    const rank = ask('Rank (e.g., C/Col)', true);
    if (rank === null) return;
    const role = ask('Role (e.g., Commander)', true);
    if (role === null) return;
    const reportsTo = ask('Reports To (optional)') || '';
    const email = ask('Email', true);
    if (email === null) return;
    const cellPhone = ask('Cell Phone (optional)') || '';
    const officePhone = ask('Office Phone (optional)') || '';
    const officeLocation = ask('Office Location (optional)') || '';

    const backendId = Config.getBackendId();
    const sheet = backendId ? SpreadsheetApp.openById(backendId).getSheetByName('Leadership Backend') : null;
    if (!sheet) {
      ui.alert('Leadership Backend sheet not found.');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    const targetRow = Math.max(3, sheet.getLastRow() + 1);
    const row: string[] = Array.from({ length: headers.length }, () => '');
    const set = (key: string, val: string) => {
      const idx = headers.indexOf(key);
      if (idx >= 0) row[idx] = val;
    };

    set('last_name', lastName);
    set('first_name', firstName);
    set('rank', rank);
    set('role', role);
    set('reports_to', reportsTo);
    set('email', email);
    set('cell_phone', cellPhone);
    set('office_phone', officePhone);
    set('office_location', officeLocation);

    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    // Sync to frontend after adding.
    try {
      SetupService.syncLeadershipBackendToFrontend();
    } catch (err) {
      Log.warn(`Unable to sync leadership to frontend after add: ${err}`);
    }
    ui.alert('Leadership entry added and synced to frontend.');
  } catch (err) {
    Log.warn(`Unable to add leadership entry: ${err}`);
  }
}

function fixAttendanceHeaders() {
  try {
    const frontendId = Config.getFrontendId();
    const ss = frontendId ? SpreadsheetApp.openById(frontendId) : null;
    const sheet = ss ? ss.getSheetByName('Attendance') : null;
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Attendance sheet not found in frontend.');
      return;
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;
    const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));
    const normalizedHeaders = headers.map((h) => h.trim().toLowerCase().replace(/\s+/g, ''));

    // Left-justify and bold all headers first.
    sheet.getRange(2, 1, 1, lastCol).setHorizontalAlignment('left').setFontWeight('bold');

    const findIdx = (name: string) => normalizedHeaders.findIndex((h) => h === name.toLowerCase().replace(/\s+/g, ''));
    const llabIdx = findIdx('LLAB');
    const overallIdx = findIdx('Overall');

    const dataRows = Math.max(1, sheet.getLastRow() - 2);
    const centerCol = (idx: number) => {
      if (idx < 0) return;
      const col = idx + 1;
      sheet.getRange(2, col, 1, 1).setHorizontalAlignment('center');
      sheet.getRange(3, col, dataRows, 1).setHorizontalAlignment('center');
    };
    centerCol(llabIdx);
    centerCol(overallIdx);

    // Set font size 5 and wrap for headers after LLAB.
    const startAfterLlab = llabIdx >= 0 ? llabIdx + 2 : 1;
    if (startAfterLlab <= lastCol) {
      const width = lastCol - startAfterLlab + 1;
      sheet.getRange(2, startAfterLlab, 1, width).setFontSize(5).setWrap(true).setHorizontalAlignment('left');
    }

    const gradientColumns = [llabIdx, overallIdx].filter((idx) => idx >= 0).map((idx) => idx + 1);

    const eventStartCol = startAfterLlab;
    const eventWidth = Math.max(0, lastCol - eventStartCol + 1);
    const eventRange = eventWidth > 0 ? sheet.getRange(3, eventStartCol, dataRows, eventWidth) : null;

    // Rebuild conditional formatting rules, removing overlaps with gradient columns (keep existing event rules/colors intact).
    const rules = sheet.getConditionalFormatRules().filter((rule) => {
      try {
        const ranges = rule.getRanges ? rule.getRanges() : [];
        return !ranges.some((rg) => {
          const colStart = rg.getColumn();
          const colEnd = colStart + rg.getNumColumns() - 1;
          const rowStart = rg.getRow();
          const rowEnd = rowStart + rg.getNumRows() - 1;

          const touchesGradient = gradientColumns.some((col) => col >= colStart && col <= colEnd);
          return touchesGradient;
        });
      } catch (err) {
        Log.warn(`Skipping rule during conditional formatting rebuild: ${err}`);
        return true;
      }
    });

    const addGradientScale = (col: number) => {
      const range = sheet.getRange(3, col, dataRows, 1);
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#e67c73', SpreadsheetApp.InterpolationType.NUMBER, '0.8')
        .setGradientMidpointWithValue('#ffce65', SpreadsheetApp.InterpolationType.NUMBER, '0.9')
        .setGradientMaxpointWithValue('#57bb8a', SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setRanges([range])
        .build();
      rules.push(rule);
    };

    gradientColumns.forEach(addGradientScale);

    // Data validation + formatting for event columns (past LLAB/Overall)
    if (eventRange && eventWidth > 0) {
      try {
        // Preserve existing validation/colors: only fill gaps; otherwise reuse rule across the event matrix.
        const validationRows = eventRange.getDataValidations();
        const existingValidation = validationRows.reduce<GoogleAppsScript.Spreadsheet.DataValidation | null>((acc, row) => {
          if (acc) return acc;
          const found = row.find((v) => v !== null) || null;
          return found || null;
        }, null);

        const hasMissingValidation = validationRows.some((row) => row.some((v) => v === null));
        if (existingValidation && hasMissingValidation) {
          const filled = validationRows.map((row) => row.map((v) => v || existingValidation));
          eventRange.setDataValidations(filled);
        } else if (!existingValidation) {
          const codesRange = ss ? ss.getRange('Data Legend!$J$3:$J') : null;
          if (codesRange) {
            const validation = SpreadsheetApp.newDataValidation()
              .requireValueInRange(codesRange, true)
              .setAllowInvalid(false)
              .setHelpText('Select attendance code')
              .build();
            eventRange.setDataValidation(validation);
          }
        }
      } catch (err) {
        Log.warn(`Unable to set attendance data validation: ${err}`);
      }

      // Center and bold all event cells to improve readability.
      eventRange.setHorizontalAlignment('center').setFontWeight('bold');
    }

    try {
      sheet.setConditionalFormatRules(rules);
    } catch (err) {
      Log.warn(`Unable to set conditional format rules for attendance sheet: ${err}`);
    }
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert('Attendance headers updated.');
  } catch (err) {
    Log.warn(`Unable to fix attendance headers: ${err}`);
  }
}

function showMenuHelp() {
  SetupService.showMenuHelp();
}

// Installable onOpen for frontend spreadsheet
function onFrontendOpen() {
  addShamrockMenu();
}

// Installable onOpen for backend spreadsheet
function onBackendOpen() {
  addShamrockMenu();
}

// Prompt-driven filler for attendance events with flexible selectors.
function fillAttendanceEventPrompt() {
  const ui = SpreadsheetApp.getUi();
  const email = ((): string => {
    try {
      return Session.getActiveUser().getEmail();
    } catch {
      return '';
    }
  })();

  const parseList = (raw: string, separators: RegExp = /[,|]/): string[] =>
    raw
      .split(separators)
      .map((s) => s.trim())
      .filter(Boolean);

  const eventResp = ui.prompt(
    'Fill attendance events',
    [
      'Select events to fill.',
      'Examples:',
      ' - all',
      ' - names:TW-17 Mando|TW-17 Secondary',
      ' - starts:TW-, ends:Secondary',
      ' - contains:Secondary',
      ' - TW-18 (unprefixed tokens treated as names)',
      'Use commas or pipes to separate multiple tokens.',
    ].join('\n'),
    ui.ButtonSet.OK_CANCEL,
  );
  if (eventResp.getSelectedButton() !== ui.Button.OK) return;
  const eventRaw = eventResp.getResponseText().trim();
  if (!eventRaw) {
    ui.alert('An event selector is required.');
    return;
  }

  const eventSelector = (() => {
    const selector = { names: [] as string[], startsWith: [] as string[], endsWith: [] as string[], contains: [] as string[], all: false };
    const tokens = eventRaw.split(/[,|;]/).map((t) => t.trim()).filter(Boolean);
    tokens.forEach((tok) => {
      const tokLower = tok.toLowerCase();
      if (tokLower === 'all' || tokLower === 'all events') {
        selector.all = true;
        return;
      }
      const [keyRaw, valRaw] = tok.includes(':') ? [tok.slice(0, tok.indexOf(':')), tok.slice(tok.indexOf(':') + 1)] : ['names', tok];
      const key = keyRaw.trim().toLowerCase();
      const val = valRaw.trim();
      if (!val && key !== 'all') return;
      if (key === 'all') selector.all = true;
      else if (key === 'names' || key === 'name') selector.names.push(...parseList(val));
      else if (key === 'starts' || key === 'starts_with') selector.startsWith.push(...parseList(val));
      else if (key === 'ends' || key === 'ends_with') selector.endsWith.push(...parseList(val));
      else if (key === 'contains') selector.contains.push(...parseList(val));
      else selector.names.push(...parseList(tok)); // fallback as name token
    });
    if (!selector.all && !selector.names.length && !selector.startsWith.length && !selector.endsWith.length && !selector.contains.length) {
      selector.names.push(eventRaw);
    }
    return selector;
  })();

  const cadetResp = ui.prompt(
    'Target cadets',
    [
      'Select cadets (union of criteria). Leave blank for all.',
      'Examples:',
      ' - all',
      ' - cadet:Doe, Jane|Smith, John',
      ' - flight:Alpha|Bravo',
      ' - university:Trine',
      ' - as_year:AS300',
      ' - abroad',
      'Combine with commas or pipes for multiple criteria.',
    ].join('\n'),
    ui.ButtonSet.OK_CANCEL,
  );
  if (cadetResp.getSelectedButton() !== ui.Button.OK) return;
  const cadetRaw = cadetResp.getResponseText().trim();

  const cadetSelector = (() => {
    const selector = { cadets: [] as string[], flights: [] as string[], universities: [] as string[], asYears: [] as string[], includeAbroad: false };
    if (!cadetRaw || cadetRaw.toLowerCase() === 'all') return selector;
    let tokens = cadetRaw.split(';').map((t) => t.trim()).filter(Boolean);
    if (tokens.length === 1) tokens = cadetRaw.split('|').map((t) => t.trim()).filter(Boolean);
    if (!tokens.length) tokens = [cadetRaw];
    tokens.forEach((tok) => {
      const [keyRaw, valRaw] = tok.includes(':') ? [tok.slice(0, tok.indexOf(':')), tok.slice(tok.indexOf(':') + 1)] : ['cadet', tok];
      const key = keyRaw.trim().toLowerCase();
      const val = valRaw.trim();
      if (!val && key !== 'abroad') return;
      if (key === 'abroad') selector.includeAbroad = true;
      else if (key === 'cadet' || key === 'name') selector.cadets.push(...parseList(val, /[|;]/));
      else if (key === 'flight') selector.flights.push(...parseList(val));
      else if (key === 'university' || key === 'uni') selector.universities.push(...parseList(val));
      else if (key === 'as_year' || key === 'asyear' || key === 'as') selector.asYears.push(...parseList(val));
      else selector.cadets.push(...parseList(tok));
    });
    return selector;
  })();

  const codeResp = ui.prompt('Attendance code', 'Enter attendance code to set (e.g., N/A, P, U, E, ED, ER):', ui.ButtonSet.OK_CANCEL);
  if (codeResp.getSelectedButton() !== ui.Button.OK) return;
  const code = codeResp.getResponseText().trim();
  if (!code) {
    ui.alert('Attendance code is required.');
    return;
  }

  try {
    const filled = AttendanceService.fillEventColumn({ eventSelector, code, cadetSelector, actorEmail: email, actorRole: 'menu_bulk_fill' });
    ui.alert(`Filled ${filled} cadet-event cells with code '${code}'.`);
  } catch (err) {
    ui.alert(`Unable to fill event: ${err}`);
  }
}

// Installable onEdit for Excusals Management spreadsheet: mirror decisions back to backend + attendance
function onExcusalsManagementEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  ExcusalsService.handleExcusalsManagementEdit(e);
}

// Installable onEdit for frontend spreadsheet: mirror allowed Directory edits back to backend with audit + propagation.
function onFrontendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e?.range?.getSheet();
  const range = e?.range;
  if (sheet && range) {
    const sheetName = sheet.getName();
    const notation = range.getA1Notation();
    const newVal = String((e as any)?.value ?? range.getValue() ?? '').substring(0, 50);
    Log.info(`[Frontend] ${sheetName} ${notation} -> "${newVal}"`);
  }
  FrontendEditService.onEdit(e);
}

// Installable onEdit for backend spreadsheet: resync directory when backend changes, handle excusals decisions.
function onBackendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  if (PauseService.isPaused()) {
    Log.info('Automation paused; skipping onBackendEdit processing.');
    return;
  }

  try {
    const sheet = e?.range?.getSheet();
    if (!sheet) return;
    const sheetName = sheet.getName();

    // Handle Excusals Backend edits (decision workflow) early and return
    if (sheetName === 'Excusals Backend') {
      ExcusalsService.handleExcusalsBackendEdit(e);
      return;
    }

    const range = e?.range;
    if (range) {
      const notation = range.getA1Notation();
      const oldVal = String((e as any)?.oldValue ?? '').substring(0, 50);
      const newVal = String((e as any)?.value ?? range.getValue() ?? '').substring(0, 50);
      Log.info(`[Backend] ${sheetName} ${notation}: "${oldVal}" -> "${newVal}"`);
    }
    try {
      const backendId = Config.getBackendId();
      if (backendId) {
        const row = e?.range?.getRow() || 0;
        const col = e?.range?.getColumn() || 0;
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
        const header = headers[col - 1] || '';
        const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
        const normalize = (v: any) => String(v || '').toLowerCase();
        let targetKey = `${sheetName}!R${row}C${col}`;
        if (sheetName === 'Directory Backend') {
          const emailIdx = headers.indexOf('email');
          const lastIdx = headers.indexOf('last_name');
          const firstIdx = headers.indexOf('first_name');
          const email = emailIdx >= 0 ? normalize(rowValues[emailIdx]) : '';
          const last = lastIdx >= 0 ? normalize(rowValues[lastIdx]) : '';
          const first = firstIdx >= 0 ? normalize(rowValues[firstIdx]) : '';
          targetKey = email || (last && first ? `${last},${first}` : targetKey);
        }

        const oldValue = String((e as any)?.oldValue ?? '');
        const newValue = String((e as any)?.value ?? e?.range?.getValue() ?? '');

        FrontendEditService.logAuditEntry({
          backendId,
          targetRange: `${sheetName}!${e?.range?.getA1Notation() || ''}`,
          targetKey,
          header,
          oldValue,
          newValue,
          targetSheet: sheetName,
          targetTable: sheetName.toLowerCase().replace(/\s+/g, '_'),
          role: 'backend_editor',
          source: 'onBackendEdit',
        });
        Log.info(`[Backend] ${targetKey} ${header} changed: \"${oldValue}\" -> \"${newValue}\"`);
      }
    } catch (err) {
      Log.warn(`Backend audit logging failed: ${err}`);
    }

    if (sheetName === 'Directory Backend') {
      SetupService.syncDirectoryFrontend();
      return;
    }

    if (sheetName === 'Data Legend') {
      SyncService.syncByBackendSheetName('Data Legend');
      SetupService.applyFrontendFormatting();
      return;
    }

    if (sheetName === 'Events Backend') {
      SetupService.refreshEventsArtifacts();
      return;
    }

    if (sheetName === 'Attendance Backend') {
      SetupService.rebuildAttendanceMatrix();
      SetupService.applyAttendanceBackendFormattingPublic();
      return;
    }

    // Sync other mapped tables when edited.
    SyncService.syncByBackendSheetName(sheetName);
  } catch (err) {
    Log.warn(`onBackendEdit failed: ${err}`);
  }
}

// Debug helper: logs current sheet headers, sizes, and form destinations.
function dumpShamrockStructure() {
  Debug.dumpShamrockStructure();
}

// Debug helper: saves structure snapshot to Drive as JSON and logs the file ID.
function dumpShamrockStructureToDrive() {
  Debug.dumpShamrockStructureToDrive();
}

// Form triggers
function onDirectoryFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
  FormHandlers.onDirectoryFormSubmit(e);
}

// Debug: replay the latest Directory form response through the handler.
function replayLatestDirectoryFormResponse() {
  const ok = DirectoryService.replayLatestDirectoryFormResponse();
  SpreadsheetApp.getUi().alert(ok ? 'Replayed latest Directory form response.' : 'No Directory form response replayed.');
}

function onAttendanceFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
  FormHandlers.onAttendanceFormSubmit(e);
}

function onExcusalsFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
  FormHandlers.onExcusalsFormSubmit(e);
}

// Time-based trigger: reconcile frontend Directory edits to backend (handles edits by unauthorized users).
function reconcilePendingDirectoryEdits() {
  if (PauseService.isPaused()) {
    Log.info('Automation paused; skipping Directory reconciliation.');
    return;
  }
  try {
    const result = FrontendEditService.reconcilePendingDirectoryEdits();
    if (result.updated > 0) {
      Log.info(`Reconciled ${result.updated} Directory edits from frontend to backend`);
      // After reconciling, sync backend -> frontend and rebuild attendance
      SyncService.syncByBackendSheetName('Directory');
      AttendanceService.rebuildMatrix();
    }
    if (result.missing > 0) {
      Log.warn(`${result.missing} frontend Directory rows not found in backend`);
    }
  } catch (err) {
    Log.warn(`reconcilePendingDirectoryEdits failed: ${err}`);
  }
}
