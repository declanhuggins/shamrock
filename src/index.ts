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
      return Session.getActiveUser().getEmail();
    } catch (err) {
      Log.warn(`Unable to read active user email for menu gate: ${err}`);
      return '';
    }
  })();

  const allowed = getAllowedMenuUsers();
  const emailLower = (email || '').toLowerCase();

  if (!allowed.includes(emailLower)) {
    Log.warn(`Menu suppressed for user=${email || 'unknown'}; not in allowed list`);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SHAMROCK')
    .addItem('Run setup (ensure-exists)', 'setup')
    .addItem('Refresh Data Legend + validations', 'refreshDataLegendAndFrontend')
    .addItem('Sync Directory Frontend', 'syncDirectoryFrontend')
    .addItem('Refresh Events + Attendance', 'refreshEventsArtifacts')
    .addItem('Rebuild Attendance Matrix', 'rebuildAttendanceMatrix')
    .addItem('Rebuild Attendance Form', 'rebuildAttendanceForm')
    .addItem('Prune Attendance Response Duplicates', 'pruneAttendanceResponseColumns')
    .addItem('Apply Frontend Formatting', 'applyFrontendFormatting')
    .addItem('Toggle Frontend Formatting (on/off)', 'toggleFrontendFormatting')
    .addItem('Reapply Frontend Protections', 'reapplyFrontendProtections')
    .addItem('Archive Core Sheets', 'archiveCoreSheets')
    .addItem('Restore Core Sheets from Archive', 'restoreCoreSheetsFromArchive')
    .addSeparator()
    .addItem('Export category (JSON)', 'exportCategory')
    .addItem('Import category (JSON)', 'importCategory')
    .addItem('Export Events CSV (backend)', 'exportEventsCsv')
    .addItem('Import Events CSV (backend)', 'importEventsCsv')
    .addItem('Export Attendance CSV (backend)', 'exportAttendanceCsv')
    .addItem('Import Attendance CSV (backend)', 'importAttendanceCsv')
    .addItem('Import Leadership CSV (backend)', 'importLeadershipCsv')
    .addItem('Import cadet CSV -> Directory Backend', 'importCadetCsv')
    .addItem('Sync Directory to Frontend', 'syncDirectoryFrontend')
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

function exportCategory() {
  AdminService.exportCategory();
}

function importCategory() {
  AdminService.importCategory();
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

function importLeadershipCsv() {
  AdminService.importCategoryCsv(undefined, 'cadre');
}

function importCadetCsv() {
  AdminService.importCadetCsv();
}

function syncDirectoryFrontend() {
  SetupService.syncDirectoryFrontend();
}

function refreshDataLegendAndFrontend() {
  SetupService.refreshDataLegendAndFrontend();
}

function refreshEventsArtifacts() {
  SetupService.refreshEventsArtifacts();
}

function rebuildAttendanceMatrix() {
  SetupService.rebuildAttendanceMatrix();
}

function rebuildAttendanceForm() {
  SetupService.rebuildAttendanceForm();
}

function applyFrontendFormatting() {
  SetupService.applyFrontendFormatting();
}

function toggleFrontendFormatting() {
  SetupService.toggleFrontendFormatting();
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

function pruneAttendanceResponseColumns() {
  SetupService.pruneAttendanceResponseColumns();
}

// Installable onOpen for frontend spreadsheet
function onFrontendOpen() {
  addShamrockMenu();
}

// Installable onOpen for backend spreadsheet
function onBackendOpen() {
  addShamrockMenu();
}

// Installable onEdit for backend spreadsheet: resync directory when backend changes.
function onBackendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  try {
    const sheet = e?.range?.getSheet();
    if (!sheet) return;
    const name = sheet.getName();
    if (name === 'Directory Backend') {
      SetupService.syncDirectoryFrontend();
      return;
    }

    if (name === 'Data Legend') {
      SyncService.syncByBackendSheetName('Data Legend');
      SetupService.applyFrontendFormatting();
      return;
    }

    if (name === 'Events Backend') {
      SetupService.refreshEventsArtifacts();
      return;
    }

    if (name === 'Attendance Backend') {
      SetupService.rebuildAttendanceMatrix();
      SetupService.applyFrontendFormatting();
      SetupService.applyAttendanceBackendFormattingPublic();
      return;
    }

    // Sync other mapped tables when edited.
    SyncService.syncByBackendSheetName(name);
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

function onAttendanceFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
  FormHandlers.onAttendanceFormSubmit(e);
}

function onExcusalFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
  FormHandlers.onExcusalFormSubmit(e);
}
