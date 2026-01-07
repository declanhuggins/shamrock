// Data Legend maintenance: populate backend legend from canonical Arrays values and sync to frontend.

namespace DataLegendService {
  interface LegendColumn {
    header: string; // sheet header key
    rangeName: string; // named range to set
    values: string[];
  }

  const LEGEND_HEADERS = [
    'as_year_options',
    'flight_options',
    'squadron_options',
    'university_options',
    'dorm_options',
    'home_state_options',
    'cip_broad_area_options',
    'afsc_options',
    'flight_path_status_options',
    'attendance_code_options',
  ];
  const LEGEND_DISPLAY_HEADERS = [
    'AS Year Options',
    'Flight Options',
    'Squadron Options',
    'University Options',
    'Dorm Options',
    'Home State Options',
    'CIP Broad Area Options',
    'AFSC Options',
    'Flight Path Status Options',
    'Attendance Code Options',
  ];

  function legendColumns(): LegendColumn[] {
    const A: any = (globalThis as any).Arrays;
    if (!A) {
      Log.warn('Arrays namespace not available; Data Legend not populated.');
      return [];
    }
    return [
      { header: 'as_year_options', rangeName: 'AS_YEARS', values: A.AS_YEARS || [] },
      { header: 'flight_options', rangeName: 'FLIGHTS', values: A.FLIGHTS || [] },
      { header: 'squadron_options', rangeName: 'SQUADRONS', values: A.SQUADRONS || [] },
      { header: 'university_options', rangeName: 'UNIVERSITIES', values: A.UNIVERSITIES || [] },
      { header: 'dorm_options', rangeName: 'DORMS', values: A.DORMS || [] },
      { header: 'home_state_options', rangeName: 'HOME_STATES', values: A.HOME_STATES || [] },
      { header: 'cip_broad_area_options', rangeName: 'CIP_BROAD_AREAS', values: A.CIP_BROAD_AREAS || [] },
      { header: 'afsc_options', rangeName: 'AFSC_OPTIONS', values: A.AFSC_OPTIONS || [] },
      { header: 'flight_path_status_options', rangeName: 'FLIGHT_PATH_STATUSES', values: A.FLIGHT_PATH_STATUSES || [] },
      { header: 'attendance_code_options', rangeName: 'ATTENDANCE_CODES', values: A.ATTENDANCE_CODES || [] },
    ];
  }

  function getBackendLegendSheet(): GoogleAppsScript.Spreadsheet.Sheet | null {
    const backendId = Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.BACKEND_SHEET_ID) || '';
    if (!backendId) return null;
    return SheetUtils.getSheet(backendId, 'Data Legend');
  }

  function ensureLegendHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const width = LEGEND_HEADERS.length;
    sheet.getRange(1, 1, 1, width).setValues([LEGEND_HEADERS]);
    sheet.getRange(2, 1, 1, width).setValues([LEGEND_DISPLAY_HEADERS]);
    // Trim extra columns if present
    const maxCols = sheet.getMaxColumns();
    if (maxCols > width) {
      sheet.deleteColumns(width + 1, maxCols - width);
    }
  }

  export function refreshLegendFromArrays() {
    const sheet = getBackendLegendSheet();
    if (!sheet) return;
    const cols = legendColumns();
    if (!cols.length) return;
    ensureLegendHeaders(sheet);
    const maxLen = cols.reduce((m, c) => Math.max(m, c.values.length), 0);
    // Clear existing data rows
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, sheet.getMaxColumns()).clearContent();
    }
    if (maxLen === 0) return;
    const matrix: string[][] = [];
    for (let i = 0; i < maxLen; i++) {
      const row: string[] = [];
      cols.forEach((c) => {
        row.push(c.values[i] || '');
      });
      matrix.push(row);
    }
    sheet.getRange(3, 1, matrix.length, cols.length).setValues(matrix);
  }
}
