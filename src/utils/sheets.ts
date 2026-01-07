// Sheet utilities for header-driven operations.

namespace SheetUtils {
  export interface TableData {
    headers: string[];
    rows: any[];
  }

  export function getSheet(spreadsheetId: string, name: string): GoogleAppsScript.Spreadsheet.Sheet | null {
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      return ss.getSheetByName(name);
    } catch (err) {
      Log.error(`Unable to open sheet ${name} in ${spreadsheetId}: ${err}`);
      return null;
    }
  }

  // Reads table data assuming row 1 = machine headers, row 2 = display headers, data starts at row 3.
  export function readTable(sheet: GoogleAppsScript.Spreadsheet.Sheet): TableData {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
    if (lastRow < 3) {
      return { headers, rows: [] };
    }
    const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const rows = values.map((row) => {
      const obj: Record<string, any> = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx];
      });
      return obj;
    });
    return { headers, rows };
  }

  // Writes table data (array of objects) starting at row 3, preserving existing headers.
  export function writeTable(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: Record<string, any>[]) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
    // Clear existing data rows (row 3 onward)
    const lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
    }
    if (!rows.length) return;
    const output = rows.map((r) => headers.map((h) => r[h] ?? ''));
    sheet.getRange(3, 1, output.length, headers.length).setValues(output);
  }

  // Appends rows to the table starting at the first empty row after header rows.
  export function appendRows(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: Record<string, any>[]) {
    if (!rows.length) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
    const startRow = Math.max(3, sheet.getLastRow() + 1);
    const output = rows.map((r) => headers.map((h) => r[h] ?? ''));
    sheet.getRange(startRow, 1, output.length, headers.length).setValues(output);
  }
}
