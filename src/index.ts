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
    .addSubMenu(
      ui
        .createMenu('Setup & Automations')
        .addItem('Run setup (ensure-exists)', 'setup')
        .addItem('Pause automations (defer sync)', 'pauseAutomations')
        .addItem('Resume automations (process pending)', 'resumeAutomations')
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
        .addItem('Prune Attendance Response Duplicates', 'pruneAttendanceResponseColumns')
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
    )
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

function pruneAttendanceResponseColumns() {
  SetupService.pruneAttendanceResponseColumns();
}

function pruneExcusalResponseColumns() {
  SetupService.pruneExcusalResponseColumns();
}

function debugExcusalResponseColumnsVerbose() {
  SetupService.debugExcusalResponseColumnsVerbose();
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

// Installable onEdit for frontend spreadsheet: mirror allowed Directory edits back to backend with audit + propagation.
function onFrontendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  FrontendEditService.onEdit(e);
}

// Installable onEdit for backend spreadsheet: resync directory when backend changes.
function onBackendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  if (PauseService.isPaused()) {
    Log.info('Automation paused; skipping onBackendEdit processing.');
    return;
  }

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
