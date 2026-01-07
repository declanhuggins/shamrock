// Frontend formatting: band tables and apply data validations from Data Legend.

namespace FrontendFormattingService {
  interface NamedRangeDef {
    name: string;
    range: GoogleAppsScript.Spreadsheet.Range;
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

  function buildNamedRanges(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): NamedRangeDef[] {
    const sheet = ss.getSheetByName('Data Legend');
    if (!sheet) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());

    const mapping: Record<string, string> = {
      as_year_options: 'AS_YEARS',
      flight_options: 'FLIGHTS',
      squadron_options: 'SQUADRONS',
      university_options: 'UNIVERSITIES',
      dorm_options: 'DORMS',
      home_state_options: 'HOME_STATES',
      cip_broad_area_options: 'CIP_BROAD_AREAS',
      afsc_options: 'AFSC_OPTIONS',
      flight_path_status_options: 'FLIGHT_PATH_STATUSES',
      attendance_code_options: 'ATTENDANCE_CODES',
    };

    const lastRow = sheet.getLastRow();
    const defs: NamedRangeDef[] = [];
    headers.forEach((header, idx) => {
      const rangeName = mapping[header];
      if (!rangeName) return;
      const col = idx + 1;
      const rowsCount = Math.max(0, lastRow - 2);
      if (rowsCount === 0) return;
      const values = sheet.getRange(3, col, rowsCount, 1).getValues().map((r) => String(r[0] || ''));
      let nonEmpty = -1;
      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i].trim() !== '') {
          nonEmpty = i;
          break;
        }
      }
      if (nonEmpty < 0) return;
      const length = nonEmpty + 1;
      const range = sheet.getRange(3, col, length, 1);
      defs.push({ name: rangeName, range });
    });
    return defs;
  }

  function applyBandingToFrontendTables(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    Schemas.FRONTEND_TABS.forEach((tab) => {
      if (tab.name === 'FAQs') return; // FAQs is freeform text, no table banding
      const sheet = ss.getSheetByName(tab.name);
      if (!sheet) return;
      const lastRow = Math.max(sheet.getLastRow(), 3);
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      sheet.getBandings().forEach((b) => b.remove());
      const bandRange = sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol);
      bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    });
  }

  function applyDirectoryValidations(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Directory');
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || ''));
    const headerIndex = (name: string) => headers.indexOf(name);

    const map: Record<string, string> = {
      as_year: 'AS_YEARS',
      flight: 'FLIGHTS',
      squadron: 'SQUADRONS',
      university: 'UNIVERSITIES',
      dorm: 'DORMS',
      home_state: 'HOME_STATES',
      cip_broad_area: 'CIP_BROAD_AREAS',
      cip_code: 'CIP_CODES',
      desired_assigned_afsc: 'AFSC_OPTIONS',
      flight_path_status: 'FLIGHT_PATH_STATUSES',
    };

    Object.entries(map).forEach(([field, rangeName]) => {
      const colIdx = headerIndex(field);
      if (colIdx < 0) return;
      const namedRange = ss.getRangeByName(rangeName);
      if (!namedRange) return;
      const dataRange = sheet.getRange(3, colIdx + 1, Math.max(1, sheet.getMaxRows() - 2), 1);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(namedRange, true)
        .setAllowInvalid(false)
        .build();
      dataRange.setDataValidation(rule);
    });
  }

  function applyAttendanceValidations(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const namedRange = ss.getRangeByName('ATTENDANCE_CODES');
    if (!namedRange) return;

    const applyToSheet = (sheetName: string) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim().toLowerCase());
      if (!headers.length) return;
      const fixed = new Set(['last_name', 'first_name', 'as_year', 'flight', 'squadron', 'llab_attendance_pct', 'overall_attendance_pct']);
      const startRow = 3;
      const numRows = Math.max(1, sheet.getMaxRows() - 2);
      headers.forEach((h, idx) => {
        if (fixed.has(h)) return;
        const col = idx + 1;
        const dataRange = sheet.getRange(startRow, col, numRows, 1);
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(namedRange, true)
          .setAllowInvalid(false)
          .build();
        dataRange.setDataValidation(rule);
      });
    };

    applyToSheet('Attendance');
    applyToSheet('Attendance Matrix Backend');
  }

  export function applyAll(frontendId: string) {
    const ss = openFrontend(frontendId);
    if (!ss) return;
    const namedRanges = buildNamedRanges(ss);
    namedRanges.forEach((def) => {
      try {
        ss.setNamedRange(def.name, def.range);
      } catch (err) {
        Log.warn(`Unable to set named range ${def.name}: ${err}`);
      }
    });

    applyDirectoryValidations(ss);
    applyAttendanceValidations(ss);

    if (shouldSkipSheetFormatting()) {
      Log.info('Sheet formatting skipped due to DISABLE_FRONTEND_FORMATTING property. Validations still applied.');
      return;
    }

    freezeTopTwoRowsAllSheets(ss);
    applyBandingToFrontendTables(ss);
    applyDirectoryFormatting(ss);
    applyLeadershipFormatting(ss);
    applyDashboardFormatting(ss);
    applyFaqsFormatting(ss);
    applyDataLegendFormatting(ss);
    applyAttendanceFormatting(ss);
  }

  function freezeTopTwoRowsAllSheets(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    ss.getSheets().forEach((sheet) => freezeTopTwoRows(sheet));
  }

  function hideMachineHeaderRow(sheet: GoogleAppsScript.Spreadsheet.Sheet | null) {
    if (!sheet) return;
    try {
      sheet.hideRows(1);
    } catch (err) {
      Log.warn(`Unable to hide row 1 on ${sheet.getName()}: ${err}`);
    }
  }

  function pruneTrailingRows(sheet: GoogleAppsScript.Spreadsheet.Sheet | null) {
    if (!sheet) return;
    const lastDataRow = Math.max(3, sheet.getLastRow());
    const maxRows = sheet.getMaxRows();
    if (maxRows > lastDataRow) {
      sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
    }
  }

  function setDefaultFont(sheet: GoogleAppsScript.Spreadsheet.Sheet | null) {
    if (!sheet) return;
    const lastRow = Math.max(2, sheet.getMaxRows());
    const lastCol = Math.max(1, sheet.getMaxColumns());
    sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Roboto').setFontSize(10);
  }

  function freezeTopTwoRows(sheet: GoogleAppsScript.Spreadsheet.Sheet | null) {
    if (!sheet) return;
    try {
      sheet.setFrozenRows(2);
    } catch (err) {
      Log.warn(`Unable to freeze rows on ${sheet.getName()}: ${err}`);
    }
  }

  function shouldSkipSheetFormatting(): boolean {
    try {
      const prop = Config.scriptProperties().getProperty('DISABLE_FRONTEND_FORMATTING');
      return String(prop || '').toLowerCase() === 'true';
    } catch (err) {
      Log.warn(`Unable to read DISABLE_FRONTEND_FORMATTING property: ${err}`);
      return false;
    }
  }

  function setColumnWidths(sheet: GoogleAppsScript.Spreadsheet.Sheet, widths: Record<string, number>) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    Object.entries(widths).forEach(([header, width]) => {
      const idx = headers.indexOf(header);
      if (idx >= 0) sheet.setColumnWidth(idx + 1, width);
    });
  }

  function setHeaderLabels(sheet: GoogleAppsScript.Spreadsheet.Sheet, mapping: Record<string, string>) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    const display = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let dirty = false;
    headers.forEach((h, idx) => {
      if (mapping[h]) {
        display[idx] = mapping[h];
        dirty = true;
      }
    });
    if (dirty) sheet.getRange(2, 1, 1, sheet.getLastColumn()).setValues([display]);
  }

  function applyDirectoryFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Directory');
    if (!sheet) return;

    setHeaderLabels(sheet, {
      as_year: 'Year',
      class_year: 'Class',
      phone_display: 'Phone Number',
      cip_code: 'CIP',
      squadron: 'Sqdn',
      flight_path_status: 'Flight Path',
      desired_assigned_afsc: 'Desired / Assigned AFSC',
    });

    setColumnWidths(sheet, {
      as_year: 100,
      class_year: 75,
      flight: 75,
      squadron: 100,
      university: 100,
      email: 175,
      phone_display: 125,
      cip_code: 75,
      dob: 100,
      flight_path_status: 125,
    });

    // Fit-to-data widths where requested.
    ['last_name', 'first_name', 'email', 'dorm', 'home_town', 'home_state', 'cip_broad_area', 'desired_assigned_afsc', 'photo_link', 'notes'].forEach((h) => {
      const idx = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf(h);
      if (idx >= 0) sheet.autoResizeColumn(idx + 1);
    });

    // Alignments
    const dataRange = sheet.getRange(3, 1, Math.max(1, sheet.getMaxRows() - 2), sheet.getLastColumn());
    dataRange.setHorizontalAlignment('left');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    const alignCenter = (key: string) => {
      const idx = headers.indexOf(key);
      if (idx >= 0) sheet.getRange(3, idx + 1, Math.max(1, sheet.getMaxRows() - 2), 1).setHorizontalAlignment('center');
    };
    alignCenter('phone_display');
    alignCenter('cip_code');
    const dobIdx = headers.indexOf('dob');
    if (dobIdx >= 0) sheet.getRange(3, dobIdx + 1, Math.max(1, sheet.getMaxRows() - 2), 1).setHorizontalAlignment('right');

    // Freeze name columns
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2);

    pruneTrailingRows(sheet);
    hideMachineHeaderRow(sheet);
    setDefaultFont(sheet);
  }

  function columnToLetter(col: number): string {
    let temp = '';
    let n = col;
    while (n > 0) {
      const rem = (n - 1) % 26;
      temp = String.fromCharCode(65 + rem) + temp;
      n = Math.floor((n - 1) / 26);
    }
    return temp;
  }

  function applyLeadershipFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Leadership');
    if (!sheet) return;
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2);

    // Hide reports_to helper column for charting.
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
    const reportsIdx = headers.indexOf('reports_to');
    if (reportsIdx >= 0) {
      sheet.hideColumns(reportsIdx + 1, 1);
    }

    pruneTrailingRows(sheet);
    hideMachineHeaderRow(sheet);
    setDefaultFont(sheet);
  }

  function applyDashboardFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Dashboard');
    if (!sheet) return;

    hideMachineHeaderRow(sheet);
    setDefaultFont(sheet);

    // Section titles
    sheet.getRange('A2').setValue('Quick Links');
    sheet.getRange('D2').setValue('Key Metrics');
    sheet.getRange('G2').setValue('Attendance Charts');
    sheet.getRange('I2').setValue('Birthdays (auto from Directory)');

    // Quick links table
    const quickLinksHeader = [['Name', 'URL']];
    const quickLinks = [
      ['Github', '=HYPERLINK("https://github.com/declanhuggins/shamrock","Open")'],
      ['Directory Form', '=HYPERLINK("https://docs.google.com/forms/d/e/1FAIpQLSfsnrADRvYm4wPFJixTVnzL37ytfRhhRrpWgwXOJiOusOHczw/viewform?usp=dialog","Open")'],
      ['Attendance Form', '=HYPERLINK("https://docs.google.com/forms/d/e/1FAIpQLSefNJ7Y87qcJ-aaN8oEfxKpZGSe5aGVKXl35ZjMOlnzHpwdzw/viewform?usp=dialog","Open")'],
      ['Excusals Form', '=HYPERLINK("https://docs.google.com/forms/d/e/1FAIpQLSd7md2rnaOEq9EceR-6nba6NeiagnnFGgYIFdDAHR7uEz1_wg/viewform?usp=dialog","Open")'],
      ['Backend sheet (admin)', '=HYPERLINK("https://docs.google.com/spreadsheets/d/13BbH2fbkSG0eyzq_M9IAtM7qHCHxRTU7m2YH2FDNR3g/edit?usp=sharing","Open")'],
    ];
    sheet.getRange(3, 1, 1, 2).setValues(quickLinksHeader).setFontWeight('bold');
    sheet.getRange(4, 1, quickLinks.length, 2).setValues(quickLinks);

    // Key metrics placeholders
    const metricsHeader = [['Metric', 'Value']];
    const metrics = [
      ['Total Cadets', '=COUNTA(Directory!A3:A)'],
      ['Pending Excusals (backend-fed later)', ''],
      ['Upcoming Events (manual)', ''],
      ['Attendance YTD (manual)', ''],
    ];
    sheet.getRange(3, 4, 1, 2).setValues(metricsHeader).setFontWeight('bold');
    sheet.getRange(4, 4, metrics.length, 2).setValues(metrics);

    // Attendance charts placeholder
    sheet.getRange('G3').setValue('Add charts here using Attendance data').setFontStyle('italic');

    // Birthdays view sourced from Directory
    const birthdayHeaders = [['Last Name', 'First Name', 'Birthday', 'Sorted', 'Display', 'Group']];
    sheet.getRange(3, 9, 1, birthdayHeaders[0].length).setValues(birthdayHeaders).setFontWeight('bold');
    sheet.getRange('I4').setFormula(
      '=LET(\n' +
      'raw, FILTER({Directory!A3:A, Directory!B3:B, Directory!M3:M}, (Directory!A3:A<>"")*(Directory!M3:M<>"")),\n' +
      'base, DATE(YEAR(TODAY()), MONTH(INDEX(raw,,3)), DAY(INDEX(raw,,3))),\n' +
      'data, HSTACK(INDEX(raw,,1), INDEX(raw,,2), INDEX(raw,,3), base),\n' +
      'sorted, SORTBY(data, TEXT(INDEX(data,,4), "MMDD"), 1),\n' +
      'l, INDEX(sorted,,1),\n' +
      'f, INDEX(sorted,,2),\n' +
      'dob, INDEX(sorted,,3),\n' +
      'sortKey, INDEX(sorted,,4),\n' +
      'display, l&", "&f&" ("&TEXT(dob, "M/D")&")",\n' +
      'week_key, WEEKNUM(sortKey, 1)+YEAR(sortKey)*1000,\n' +
      'grp, XMATCH(week_key, UNIQUE(week_key)),\n' +
      'HSTACK(l, f, dob, TEXT(sortKey, "MM/DD"), display, grp)\n' +
      ')'
    );

    // Layout polish
    sheet.setFrozenRows(2);
    sheet.autoResizeColumns(1, 2); // quick links
    sheet.setColumnWidth(2, 160);
    sheet.autoResizeColumns(4, 2); // metrics
    sheet.setColumnWidth(7, 220); // charts note area
    sheet.setColumnWidths(9, 6, 140); // birthdays block
    sheet.getRange('I4:I').setHorizontalAlignment('left');
    sheet.getRange('J4:J').setHorizontalAlignment('left');
    sheet.getRange('K4:K').setNumberFormat('M/D/YYYY');
    sheet.getRange('L4:L').setNumberFormat('MM/DD');
    sheet.getRange('M4:M').setWrap(true);
    sheet.getRange('N4:N').setHorizontalAlignment('center');
  }

  function applyFaqsFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('FAQs');
    if (!sheet) return;

    // Remove old header rows if they match the legacy schema.
    try {
      const first = String(sheet.getRange(1, 1).getValue() || '').trim().toLowerCase();
      const second = String(sheet.getRange(2, 1).getValue() || '').trim().toLowerCase();
      if (first === 'faq' && (!second || second === 'faq')) {
        const deleteCount = Math.min(2, sheet.getMaxRows());
        if (deleteCount > 0) sheet.deleteRows(1, deleteCount);
      }
    } catch (err) {
      Log.warn(`Unable to normalize FAQ headers: ${err}`);
    }

    // Strip banding and keep a single wide column for freeform text.
    sheet.getBandings().forEach((b) => b.remove());
    const maxCols = sheet.getMaxColumns();
    if (maxCols > 1) {
      sheet.deleteColumns(2, maxCols - 1);
    }

    // Ensure no frozen rows/columns so the page behaves like a blank canvas.
    try {
      sheet.setFrozenRows(0);
      sheet.setFrozenColumns(0);
    } catch (err) {
      Log.warn(`Unable to unfreeze FAQ sheet: ${err}`);
    }

    setDefaultFont(sheet);
    sheet.setColumnWidth(1, 1000);

    // Allow freeform text starting at A1.
    const totalRows = Math.max(1, sheet.getMaxRows());
    const contentRange = sheet.getRange(1, 1, totalRows, 1);
    contentRange.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
  }

  function applyDataLegendFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Data Legend');
    if (!sheet) return;
    pruneTrailingRows(sheet);
    hideMachineHeaderRow(sheet);
    setDefaultFont(sheet);
  }

  function applyAttendanceFormatting(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return;

    // Hide machine headers, prune, defaults
    pruneTrailingRows(sheet);
    hideMachineHeaderRow(sheet);
    setDefaultFont(sheet);

    // Rename percentage headers
    const display = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim().toLowerCase());
    const llabIdx = headers.indexOf('llab_attendance_pct');
    const overallIdx = headers.indexOf('overall_attendance_pct');
    if (llabIdx >= 0) display[llabIdx] = 'LLAB';
    if (overallIdx >= 0) display[overallIdx] = 'Overall';
    sheet.getRange(2, 1, 1, sheet.getLastColumn()).setValues([display]);

    // Widths and hides
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);
    sheet.hideColumns(3, 3);
    if (sheet.getLastColumn() >= 8) {
      sheet.setColumnWidths(8, sheet.getLastColumn() - 7, 75);
    }

    // Header styling for event columns H+
    if (sheet.getLastColumn() >= 8) {
      const headerRange = sheet.getRange(2, 8, 1, sheet.getLastColumn() - 7);
      headerRange.setFontSize(5).setWrap(true);
    }

    // Alignments
    const dataRows = Math.max(1, sheet.getMaxRows() - 2);
    sheet.getRange(3, 1, dataRows, sheet.getLastColumn()).setHorizontalAlignment('left');
    if (sheet.getLastColumn() >= 8) {
      sheet.getRange(3, 8, dataRows, sheet.getLastColumn() - 7).setHorizontalAlignment('center');
    }

    // Percentage formats
    const formatPercent = (idx: number) => {
      if (idx >= 0) {
        const range = sheet.getRange(3, idx + 1, dataRows, 1);
        range.setNumberFormat('0.0%');
        range.setHorizontalAlignment('center');
      }
    };
    formatPercent(llabIdx);
    formatPercent(overallIdx);

    // Freeze first two columns
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2);

    const rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = [];

    // Attendance percentage gradient (LLAB/Overall: columns F and G)
    const maxRows = Math.max(3, sheet.getMaxRows());
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('0.8', SpreadsheetApp.InterpolationType.NUMBER, '#e67c73')
        .setGradientMidpointWithValue('0.9', SpreadsheetApp.InterpolationType.NUMBER, '#ffce65')
        .setGradientMaxpointWithValue('1', SpreadsheetApp.InterpolationType.NUMBER, '#57bb8a')
        .setRanges([sheet.getRange(1, 6, maxRows, 2)])
        .build(),
    );

    // Attendance code colors for event columns
    const startRow = 3;
    const lastRow = Math.max(startRow, sheet.getLastRow());
    const rowCount = Math.max(1, lastRow - startRow + 1);
    const codePalette: Record<string, string> = {
      P: '#C8E6C9',
      E: '#BBDEFB',
      ES: '#E1BEE7',
      ER: '#FFF9C4',
      ED: '#FFE0B2',
      T: '#E0E0E0',
      U: '#FFCDD2',
      UR: '#F8BBD0',
      MU: '#D1C4E9',
      MRS: '#C5CAE9',
      'N/A': '#F5F5F5',
      '': '#FFFFFF',
    };
    headers.forEach((h, idx) => {
      if (idx < 7) return; // event columns only
      const colLetter = columnToLetter(idx + 1);
      const colRange = sheet.getRange(startRow, idx + 1, rowCount, 1);
      Object.entries(codePalette).forEach(([code, color]) => {
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=$${colLetter}${startRow}="${code}"`)
            .setBackground(color)
            .setRanges([colRange])
            .build(),
        );
      });
    });

    try {
      sheet.setConditionalFormatRules(rules);
    } catch (err) {
      Log.warn(`Unable to set conditional formatting on Attendance: ${err}`);
    }
  }
}
