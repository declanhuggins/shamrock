// Frontend protections: lock headers, key columns, and scoped editors.

namespace ProtectionService {
  interface ProtectionOptions {
    warningOnly?: boolean;
    editors?: string[];
  }

  function openFrontend(frontendId: string): GoogleAppsScript.Spreadsheet.Spreadsheet | null {
    if (!frontendId) return null;
    try {
      return SpreadsheetApp.openById(frontendId);
    } catch (err) {
      Log.warn(`Unable to open frontend spreadsheet ${frontendId}: ${err}`);
      return null;
    }
  }

  function normalizeEditors(editors: string[]): string[] {
    return Array.from(new Set(editors.map((e) => (e || '').trim()).filter(Boolean)));
  }

  function ensureRangeProtection(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    range: GoogleAppsScript.Spreadsheet.Range,
    description: string,
    opts: ProtectionOptions = {},
  ) {
    // Remove any prior protection with the same description to avoid stacking duplicates.
    sheet
      .getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter((p) => p.getDescription && p.getDescription() === description)
      .forEach((p) => p.remove());

    const protection = range.protect().setDescription(description);
    protection.setWarningOnly(Boolean(opts.warningOnly));
    const editors = normalizeEditors(opts.editors || []);
    if (!protection.isWarningOnly() && editors.length) {
      try {
        protection.addEditors(editors);
      } catch (err) {
        Log.warn(`Unable to add editors to protection ${description}: ${err}`);
      }
    }
    return protection;
  }

  function protectFirstTwoRows(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    ss.getSheets().forEach((sheet) => {
      const name = sheet.getName();
      if (name === 'FAQs' || name === 'Dashboard') return; // handled separately
      const lastCol = Math.max(1, sheet.getLastColumn(), sheet.getMaxColumns());
      const range = sheet.getRange(1, 1, 2, lastCol);
      ensureRangeProtection(sheet, range, `${sheet.getName()}:header_rows`, { warningOnly: false });
    });
  }

  function protectFaqs(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('FAQs');
    if (!sheet) return;
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    ensureRangeProtection(sheet, range, 'FAQs:all', { warningOnly: false });
  }

  function protectDashboard(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Dashboard');
    if (!sheet) return;
    const lastRow = Math.max(3, sheet.getMaxRows());
    // Protect birthdays block (headers + data) in columns I:M
    const birthdayRange = sheet.getRange(3, 9, lastRow - 2, 5);
    ensureRangeProtection(sheet, birthdayRange, 'Dashboard:birthdays', { warningOnly: false });
  }

  function getLeadershipEmails(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string[] {
    const sheet = ss.getSheetByName('Leadership');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    const emailCol = 5; // role, rank, last, first, email
    const values = sheet.getRange(3, emailCol, lastRow - 2, 1).getValues().map((r) => String(r[0] || '').trim());
    return normalizeEditors(values);
  }

  function protectLeadership(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Leadership');
    if (!sheet) return;
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    ensureRangeProtection(sheet, range, 'Leadership:all', { warningOnly: false });
  }

  function protectDataLegend(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Data Legend');
    if (!sheet) return;
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    ensureRangeProtection(sheet, range, 'Data Legend:all', { warningOnly: false });
  }

  function protectDirectory(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Directory');
    if (!sheet) return;
    // Clear any prior sheet-level protections to avoid overlapping "except" scopes.
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
    const lastRow = Math.max(3, sheet.getMaxRows());
    const lastCol = Math.max(1, sheet.getMaxColumns());

    // Lock last/first name columns for data rows only (row 3+).
    const dataRowCount = Math.max(1, lastRow - 2);
    const nameRange = sheet.getRange(3, 1, dataRowCount, 2);
    ensureRangeProtection(sheet, nameRange, 'Directory:last_first_locked', { warningOnly: false });

    // Warn-only on header rows across visible columns (A1:S2 or to last column if narrower).
    const warnCols = Math.min(lastCol, 19); // Column S = 19
    const warnRange = sheet.getRange(1, 1, dataRowCount, warnCols);
    ensureRangeProtection(sheet, warnRange, 'Directory:warn_rest', { warningOnly: true });
  }

  function protectAttendance(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return;
    // Clear any prior sheet protections to prevent stale "except" scopes.
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
    const lastRow = Math.max(3, sheet.getMaxRows());
    const lastCol = Math.max(1, sheet.getLastColumn(), sheet.getMaxColumns());

    // Lock columns A:G (Last Name through LLAB) across all rows.
    const fixedRange = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 7));
    ensureRangeProtection(sheet, fixedRange, 'Attendance:fixed_cols', { warningOnly: false });

    // Event columns (H+): protect rows 3+ but allow leadership emails to edit.
    const eventsStartCol = 8;
    if (lastCol >= eventsStartCol) {
      const editors = getLeadershipEmails(ss);
      editors.push(Session.getEffectiveUser().getEmail());
      const eventsRange = sheet.getRange(3, eventsStartCol, lastRow - 2, lastCol - eventsStartCol + 1);
      ensureRangeProtection(sheet, eventsRange, 'Attendance:event_cols_with_leadership', {
        warningOnly: false,
        editors,
      });
    }
  }

  export function applyFrontendProtections(frontendId: string) {
    const ss = openFrontend(frontendId);
    if (!ss) return;

    protectFirstTwoRows(ss);
    protectFaqs(ss);
    protectDashboard(ss);
    protectLeadership(ss);
    protectDataLegend(ss);
    protectDirectory(ss);
    protectAttendance(ss);
  }
}