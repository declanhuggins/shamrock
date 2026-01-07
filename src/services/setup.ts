// Setup service: idempotent provisioning of spreadsheets, sheets, and forms.

namespace SetupService {
  function extractFormIdFromUrl(url: string): string | null {
    if (!url) return null;
    // Common Forms URL formats:
    // - https://docs.google.com/forms/d/e/<ID>/viewform
    // - https://docs.google.com/forms/d/<ID>/edit
    const match = url.match(/\/forms\/d\/(?:e\/)?([a-zA-Z0-9_-]+)/);
    return match?.[1] || null;
  }

  function getFormDestinationSpreadsheetId(form: GoogleAppsScript.Forms.Form): string | null {
    try {
      const anyForm = form as any;
      const destinationType = anyForm.getDestinationType?.();
      const destinationId = anyForm.getDestinationId?.();
      if (destinationType === FormApp.DestinationType.SPREADSHEET && typeof destinationId === 'string') {
        return destinationId;
      }
      return null;
    } catch {
      return null;
    }
  }

  function ensureSpreadsheet(role: Types.WorkbookRole, name: string, propertyKey: string): Types.EnsureSpreadsheetResult {
    Log.info(`Ensuring spreadsheet for role=${role}`);
    const props = Config.scriptProperties();
    const existingId = props.getProperty(propertyKey);
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;
    let created = false;

    if (existingId) {
      try {
        spreadsheet = SpreadsheetApp.openById(existingId);
        Log.info(`Found existing spreadsheet id=${existingId}`);
      } catch (err) {
        Log.warn(`Stored spreadsheet id invalid for key=${propertyKey}; creating new. Error: ${err}`);
      }
    }

    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.create(name);
      props.setProperty(propertyKey, spreadsheet.getId());
      created = true;
      Log.info(`Created spreadsheet name=${name} id=${spreadsheet.getId()}`);
    }

    return {
      role,
      id: spreadsheet.getId(),
      name: spreadsheet.getName(),
      created,
      url: spreadsheet.getUrl(),
    };
  }

  function ensureSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, schema: Types.TabSchema): Types.EnsureSheetResult {
    const { name, machineHeaders, displayHeaders } = schema;
    Log.info(`Ensuring sheet name=${name} in spreadsheet=${spreadsheet.getId()}`);
    const existingSheet = spreadsheet.getSheetByName(name);
    let sheet = existingSheet;
    let created = false;
    let headersApplied = false;

    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
      created = true;
      Log.info(`Created sheet name=${name}`);
    }

    if (sheet && machineHeaders && machineHeaders.length > 0) {
      const headerWidth = machineHeaders.length;
      const firstRow = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
      const secondRow = sheet.getRange(2, 1, 1, headerWidth).getValues()[0];
      const firstRowEmpty = firstRow.every((cell) => cell === '' || cell === null);
      const secondRowEmpty = secondRow.every((cell) => cell === '' || cell === null);

      if (firstRowEmpty) {
        sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
        headersApplied = true;
        Log.info(`Applied machine headers for ${name}`);
      }

      if (secondRowEmpty) {
        const display =
          displayHeaders && displayHeaders.length === machineHeaders.length
            ? displayHeaders
            : Headers.humanizeHeaders(machineHeaders);
        sheet.getRange(2, 1, 1, headerWidth).setValues([display]);
        headersApplied = true;
        Log.info(`Applied display headers for ${name}`);
      }

      // Trim unused columns to the exact header width to avoid excess blank space.
      const maxCols = sheet.getMaxColumns();
      if (maxCols > headerWidth) {
        const deleteCount = maxCols - headerWidth;
        Log.info(`Deleting ${deleteCount} extra columns in ${name} (keeps ${headerWidth})`);
        sheet.deleteColumns(headerWidth + 1, deleteCount);
      }
    }

    return {
      spreadsheetId: spreadsheet.getId(),
      sheetName: name,
      created,
      headersApplied,
    };
  }

  function resetSheetToSchema(sheet: GoogleAppsScript.Spreadsheet.Sheet, schema: Types.TabSchema) {
    const { machineHeaders, displayHeaders } = schema;
    if (!machineHeaders || machineHeaders.length === 0) return;
    const headerWidth = machineHeaders.length;

    // Ensure column count matches schema width.
    const maxCols = sheet.getMaxColumns();
    if (maxCols < headerWidth) {
      sheet.insertColumnsAfter(maxCols, headerWidth - maxCols);
    } else if (maxCols > headerWidth) {
      sheet.deleteColumns(headerWidth + 1, maxCols - headerWidth);
    }

    // Clear all content and reapply headers.
    sheet.clear();
    sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
    const display = displayHeaders && displayHeaders.length === headerWidth ? displayHeaders : Headers.humanizeHeaders(machineHeaders);
    sheet.getRange(2, 1, 1, headerWidth).setValues([display]);
  }

  function ensureTableForSheet(spreadsheetId: string, sheetName: string, tableId: string) {
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const sheetId = sheet.getSheetId();

      const svc = (Sheets as any)?.Spreadsheets;
      if (!svc || !svc.batchUpdate) {
        Log.warn('Sheets advanced service unavailable; cannot create tables');
        return;
      }

      const headerRow = 2; // display headers live on row 2
      const headerValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colCount = headerValues.length;
      if (colCount === 0) return;
      const endColIndex = colCount; // zero-based exclusive
      const endRowIndex = Math.max(headerRow + 1, sheet.getLastRow());

      const columnProperties = headerValues.map((name, idx) => ({
        columnIndex: idx,
        columnName: String(name || `Column ${idx + 1}`),
      }));

      // Attempt to replace any existing table with the same id.
      try {
        svc.batchUpdate({ requests: [{ deleteTable: { tableId } } as any] }, spreadsheetId);
      } catch (err) {
        Log.info(`No existing table to delete for ${tableId} on ${sheetName}: ${err}`);
      }

      svc.batchUpdate(
        {
          requests: [
            {
              addTable: {
                table: {
                  name: tableId,
                  tableId,
                  range: {
                    sheetId,
                    startColumnIndex: 0,
                    endColumnIndex: endColIndex,
                    startRowIndex: headerRow - 1, // zero-based (row 2)
                    endRowIndex,
                  },
                  columnProperties,
                },
              },
            } as any,
          ],
        },
        spreadsheetId,
      );
      Log.info(`Ensured table ${tableId} on sheet ${sheetName}`);
    } catch (err) {
      Log.warn(`Unable to ensure table ${tableId} on sheet ${sheetName}: ${err}`);
    }
  }

  function ensureResponseSheetName(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, desiredName: string): boolean {
    const current = spreadsheet.getSheetByName(desiredName);
    const candidates = spreadsheet.getSheets().filter((s) => /^Form Responses/i.test(s.getName()));

    if (current) {
      // Desired sheet already present; do not delete other response sheets to avoid breaking links.
      return true;
    }

    if (candidates.length === 0) {
      Log.info(`No response sheet found to rename to ${desiredName} in spreadsheet=${spreadsheet.getId()}`);
      return false;
    }

    const primary = candidates[0];
    if (primary.getName() !== desiredName) {
      Log.info(`Renaming response sheet ${primary.getName()} -> ${desiredName}`);
      primary.setName(desiredName);
    }

    // Leave any additional Form Responses sheets untouched to avoid deleting linked sheets; log for awareness.
    candidates.slice(1).forEach((s) => {
      if (s.getName() !== desiredName) {
        Log.warn(`Additional response sheet present (${s.getName()}); leaving as-is to avoid unlinking forms.`);
      }
    });
    return true;
  }

  function copySheetToArchive(
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
    source: GoogleAppsScript.Spreadsheet.Sheet,
    archivePrefix: string,
  ): GoogleAppsScript.Spreadsheet.Sheet | null {
    const archiveName = `${archivePrefix}${source.getName()}`.trim();

    // Replace only the canonical archive sheet for this source; leave any user-renamed archives intact.
    const existingArchive = ss.getSheetByName(archiveName);
    if (existingArchive) {
      try {
        ss.deleteSheet(existingArchive);
      } catch (err) {
        Log.warn(`Unable to delete existing archive sheet ${archiveName}: ${err}`);
        return null;
      }
    }

    let archived: GoogleAppsScript.Spreadsheet.Sheet;
    try {
      archived = source.copyTo(ss);
    } catch (err) {
      Log.warn(`Unable to copy sheet ${source.getName()} to archive ${archiveName}: ${err}`);
      return null;
    }

    try {
      archived.setName(archiveName);
    } catch (err) {
      Log.warn(`Unable to rename archive copy to ${archiveName}: ${err}`);
    }

    // Sever links: strip formulas, named ranges, and protections.
    const range = archived.getDataRange();
    range.copyTo(range, { contentsOnly: true });
    archived.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach((p) => p.remove());
    archived.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
    try {
      const protection = archived.protect().setDescription(`${archiveName} (locked)`);
      protection.setWarningOnly(false);
      try {
        protection.removeEditors(protection.getEditors());
      } catch (err) {
        Log.warn(`Unable to remove editors from ${archiveName}: ${err}`);
      }
      if (protection.canDomainEdit && protection.canDomainEdit()) {
        try {
          protection.setDomainEdit(false);
        } catch (err) {
          Log.warn(`Unable to disable domain edit on ${archiveName}: ${err}`);
        }
      }
    } catch (err) {
      Log.warn(`Unable to protect archive sheet ${archiveName}: ${err}`);
    }

    ss.setActiveSheet(archived);
    ss.moveActiveSheet(ss.getSheets().length);

    return archived;
  }

  function archiveAndResetSheets(
    spreadsheetId: string,
    schemas: Types.TabSchema[],
    names: string[],
    archivePrefix = 'Archive ',
  ) {
    if (!spreadsheetId) return;
    const ss = SpreadsheetApp.openById(spreadsheetId);

    names.forEach((name) => {
      const schema = schemas.find((s) => s.name === name);
      if (!schema || !schema.machineHeaders) return;
      const sheet = ss.getSheetByName(name);
      if (!sheet) return;

      copySheetToArchive(ss, sheet, archivePrefix);

      resetSheetToSchema(sheet, schema);
    });
  }

  function restoreFromArchiveSheets(
    spreadsheetId: string,
    schemas: Types.TabSchema[],
    names: string[],
    archivePrefix = 'Archive ',
  ) {
    if (!spreadsheetId) return;
    const ss = SpreadsheetApp.openById(spreadsheetId);

    names.forEach((name) => {
      const schema = schemas.find((s) => s.name === name);
      if (!schema || !schema.machineHeaders) return;
      let target = ss.getSheetByName(name);
        const archive = ss.getSheetByName(`${archivePrefix}${name}`);
        if (!archive) return;
      if (!target) {
        target = ss.insertSheet(name);
      }

      const values = archive.getDataRange().getValues();
      const width = Math.max(schema.machineHeaders.length, values[0]?.length || 0);

      const maxCols = target.getMaxColumns();
      if (maxCols < width) target.insertColumnsAfter(maxCols, width - maxCols);
      if (maxCols > width) target.deleteColumns(width + 1, maxCols - width);

      target.clear();
      if (values.length && values[0].length) {
        target.getRange(1, 1, values.length, values[0].length).setValues(values);
      }
    });
  }

  function ensureResponseSheetNameWithRetry(spreadsheetId: string, desiredName: string, retries = 3, delayMs = 500) {
    for (let attempt = 0; attempt < retries; attempt++) {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const ok = ensureResponseSheetName(ss, desiredName);
      if (ok) return;
      Utilities.sleep(delayMs);
    }
    Log.warn(`Unable to find response sheet for ${desiredName} after ${retries} attempts in spreadsheet=${spreadsheetId}`);
  }

  function slimAttendanceResponseSheet() {
    const { backendId } = getIds();
    if (!backendId) return;
    const ss = SpreadsheetApp.openById(backendId);
    const sheet = ss.getSheetByName(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));
    const dataRows = Math.max(0, lastRow - 1);
    const startRow = 2;

    // Headers that should not exist (legacy items we removed from the form).
    const bannedHeaders = new Set(['Submitted By Email']);

    // Group columns by header; keep the first occurrence, merge data from later duplicates.
    const indicesByHeader = new Map<string, number[]>();
    headers.forEach((h, idx) => {
      const key = h.trim();
      if (!key) return;
      const arr = indicesByHeader.get(key) || [];
      arr.push(idx + 1); // 1-based
      indicesByHeader.set(key, arr);
    });

    indicesByHeader.forEach((cols, header) => {
      if (cols.length <= 1) return;
      if (bannedHeaders.has(header)) {
        // Delete all occurrences of banned headers.
        cols
          .slice()
          .sort((a, b) => b - a)
          .forEach((col) => {
            const currentMax = sheet.getMaxColumns();
            if (col > currentMax) return;
            try {
              sheet.deleteColumn(col);
            } catch (err) {
              try {
                sheet.hideColumn(sheet.getRange(1, col));
              } catch (err2) {
                Log.warn(
                  `Unable to delete or hide banned header '${header}' column ${col} in ${Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET}: ${err}; hide failed: ${err2}`,
                );
              }
            }
          });
        return;
      }
      // Merge all duplicate columns' data together (deduping values) and write the merged value into every duplicate column.
      if (dataRows > 0) {
        const colValues = cols.map((col) => sheet.getRange(startRow, col, dataRows, 1).getValues());
        const merged: string[][] = Array.from({ length: dataRows }, () => ['']);

        for (let r = 0; r < dataRows; r++) {
          const seen = new Set<string>();
          const parts: string[] = [];
          colValues.forEach((vals) => {
            const raw = String(vals[r][0] || '').trim();
            if (!raw) return;
            raw.split('|').forEach((p) => {
              const part = p.trim();
              if (!part) return;
              if (seen.has(part)) return;
              seen.add(part);
              parts.push(part);
            });
          });
          merged[r][0] = parts.join(' | ');
        }

        cols.forEach((col) => {
          sheet.getRange(startRow, col, dataRows, 1).setValues(merged);
        });
      }

      // Attempt to delete all duplicates; the column that cannot be deleted (form-linked) will remain.
      let survivor: number | null = null;
      const sorted = cols.slice().sort((a, b) => b - a); // delete right-to-left to reduce shifting issues
      sorted.forEach((col, idx) => {
        // If we have no survivor yet and this is the last column, keep it to guarantee one remains.
        if (survivor === null && idx === sorted.length - 1) {
          survivor = col;
          return;
        }

        const currentMax = sheet.getMaxColumns();
        if (col > currentMax) return;
        try {
          sheet.deleteColumn(col);
        } catch (err) {
          // Likely the form-linked column; keep it but continue pruning other duplicates.
          survivor = survivor ?? col;
        }
      });
    });
  }

  function pruneAttendanceResponseColumnsExplicit() {
    const { backendId } = getIds();
    if (!backendId) return;
    const ss = SpreadsheetApp.openById(backendId);
    const sheet = ss.getSheetByName(Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET);
    if (!sheet) return;

    // First merge any duplicate data so deletes do not drop content.
    slimAttendanceResponseSheet();

    // Re-run pruning a few times to tolerate column shifting or prior delete failures.
    const bannedHeaders = new Set(['Submitted By Email']);
    for (let attempt = 0; attempt < 5; attempt++) {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || ''));

      const indicesByHeader = new Map<string, number[]>();
      headers.forEach((h, idx) => {
        const key = h.trim();
        if (!key) return;
        const arr = indicesByHeader.get(key) || [];
        arr.push(idx + 1);
        indicesByHeader.set(key, arr);
      });

      let changed = false;
      let sawDuplicate = false;

      indicesByHeader.forEach((cols, header) => {
        if (cols.length <= 1) return;
        if (bannedHeaders.has(header)) {
          cols
            .slice()
            .sort((a, b) => b - a)
            .forEach((col) => {
              const currentMax = sheet.getMaxColumns();
              if (col > currentMax) return;
              try {
                sheet.deleteColumn(col);
                changed = true;
              } catch (err) {
                try {
                  sheet.hideColumn(sheet.getRange(1, col));
                  changed = true;
                } catch (err2) {
                  Log.warn(
                    `Unable to delete or hide banned header '${header}' column ${col} in ${Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET}: ${err}; hide failed: ${err2}`,
                  );
                }
              }
            });
          return;
        }
        sawDuplicate = true;

        let kept = false;
        const sorted = cols.slice().sort((a, b) => b - a);
        sorted.forEach((col, idx) => {
          const remaining = sorted.length - idx;

          // Always leave at least one column untouched (last remaining if none kept yet).
          if (!kept && remaining === 1) {
            kept = true;
            return;
          }

          const currentMax = sheet.getMaxColumns();
          if (col > currentMax) return;
          try {
            sheet.deleteColumn(col);
            changed = true;
          } catch (err) {
            // Likely form-linked; keep it and continue.
            kept = true;
          }
        });
      });

      if (!sawDuplicate || !changed) break;
    }

    normalizeAttendanceBackendHeaders();
  }

  function normalizeAttendanceBackendHeaders() {
    const { backendId } = getIds();
    if (!backendId) return;
    const ss = SpreadsheetApp.openById(backendId);
    const sheet = ss.getSheetByName('Attendance Backend');
    if (!sheet) return;

    const attendanceSchema = Schemas.BACKEND_TABS.find((t) => t.name === 'Attendance Backend');
    const targetHeaders = attendanceSchema?.machineHeaders || ['submission_id', 'submitted_at', 'event', 'email', 'name', 'flight', 'cadets'];
    const displayHeaders = Headers.humanizeHeaders(targetHeaders);

    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastCol = Math.max(sheet.getLastColumn(), targetHeaders.length);
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    const sourceHeaders = (values[0] || []).map((h) => String(h || '').trim());
    const sourceLookup = new Map<string, number>();
    sourceHeaders.forEach((h, idx) => {
      const key = h.toLowerCase();
      if (key) sourceLookup.set(key, idx);
    });

    const altKeys: Record<string, string[]> = {
      submission_id: ['submission id'],
      submitted_at: ['submitted at', 'timestamp', 'submission time'],
      event: ['event'],
      email: ['email', 'email address', 'submitted by email'],
      name: ['name', 'submitted by name'],
      flight: ['flight', 'flight / crosstown (mando)', 'flight (mando pt)', 'flight / crosstown', 'flight / crosstown (llab)', 'flight (llab)'],
      cadets: ['cadets', 'cadet selections', 'cadet list'],
    };

    const headerMatches = targetHeaders.map((h) => {
      const key = h.toLowerCase();
      if (sourceLookup.has(key)) return sourceLookup.get(key)!;
      const alts = altKeys[h] || [];
      for (const alt of alts) {
        const altIdx = sourceLookup.get(alt.toLowerCase());
        if (altIdx !== undefined) return altIdx;
      }
      return -1;
    });

    // Detect if row 2 is a display/header row to skip when rebuilding data.
    const displayRowMatches = (values[1] || []).every((cell: any, idx: number) => {
      const expected = displayHeaders[idx] || '';
      return String(cell || '').trim().toLowerCase() === expected.toLowerCase();
    });
    const dataStart = displayRowMatches ? 3 : 2;
    const dataRows: any[][] = [];
    for (let r = dataStart - 1; r < lastRow; r++) {
      const row = values[r] || [];
      const out = targetHeaders.map((_, idx) => {
        const srcIdx = headerMatches[idx];
        return srcIdx >= 0 ? row[srcIdx] || '' : '';
      });
      if (out.some((v) => v !== '')) dataRows.push(out);
    }

    sheet.clear();
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    sheet.getRange(2, 1, 1, targetHeaders.length).setValues([displayHeaders]);
    if (dataRows.length) {
      sheet.getRange(3, 1, dataRows.length, targetHeaders.length).setValues(dataRows);
    }
  }

    function applyAttendanceBackendFormatting() {
      const { backendId } = getIds();
      if (!backendId) return;
      const ss = SpreadsheetApp.openById(backendId);
      const sheet = ss.getSheetByName('Attendance Backend');
      if (!sheet) return;

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
      const flightCol = headers.indexOf('flight') + 1;
      if (flightCol <= 0) return;

      const startRow = 3; // data starts after header rows
      const lastRow = Math.max(startRow, sheet.getLastRow());
      const numRows = Math.max(1, lastRow - startRow + 1);
      const dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());

      const columnToLetter = (col: number) => {
        let temp = '';
        let n = col;
        while (n > 0) {
          const rem = (n - 1) % 26;
          temp = String.fromCharCode(65 + rem) + temp;
          n = Math.floor((n - 1) / 26);
        }
        return temp;
      };

      // Clear existing rules to avoid duplicates.
      sheet.clearConditionalFormatRules();

      const palette: Record<string, string> = {
        Alpha: '#E3F2FD',
        Bravo: '#FCE4EC',
        Charlie: '#F3E5F5',
        Delta: '#E8F5E9',
        Echo: '#FFF3E0',
        Foxtrot: '#E0F7FA',
        Abroad: '#ECEFF1',
        Trine: '#FFFDE7',
        Valparaiso: '#EDE7F6',
      };

      const flightColLetter = columnToLetter(flightCol);
      const rules = Object.entries(palette).map(([flight, color]) =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=$${flightColLetter}${startRow}="${flight}"`)
          .setBackground(color)
          .setRanges([dataRange])
          .build(),
      );

      try {
        sheet.setConditionalFormatRules(rules);
      } catch (err) {
        Log.warn(`Unable to set conditional formatting on Attendance Backend: ${err}`);
      }
    }

  function getIds() {
    return {
      backendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.BACKEND_SHEET_ID) || '',
      frontendId: Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.FRONTEND_SHEET_ID) || '',
    };
  }

  function ensureFormTrigger(handlerName: string, formId: string) {
    if (!formId) {
      Log.warn(`Cannot create form trigger for handler=${handlerName}: formId missing`);
      return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    const matching = triggers.filter((t) => t.getHandlerFunction() === handlerName);
    const alreadyCorrect = matching.some((t) => {
      try {
        return t.getTriggerSourceId && t.getTriggerSourceId() === formId;
      } catch {
        return false;
      }
    });
    if (alreadyCorrect) return;

    // Clean up stale triggers for the same handler so we don't keep firing against old/deleted forms.
    matching.forEach((t) => {
      try {
        const sourceId = t.getTriggerSourceId?.();
        if (sourceId && sourceId !== formId) {
          Log.warn(`Deleting stale trigger for handler=${handlerName} sourceId=${sourceId}`);
          ScriptApp.deleteTrigger(t);
        }
      } catch {
        // Ignore; we'll create a new correct trigger below.
      }
    });

    Log.info(`Creating form submit trigger for handler=${handlerName} formId=${formId}`);
    ScriptApp.newTrigger(handlerName).forForm(formId).onFormSubmit().create();
  }

  function normalizeResponseSheetsForForms(
    spreadsheetId: string,
    forms: Array<{ formId: string; desiredSheetName: string }>,
  ) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const desiredByFormId = new Map(forms.map((f) => [f.formId, f.desiredSheetName] as const));

    const responseishSheets = ss
      .getSheets()
      .filter((s) => /^Form Responses/i.test(s.getName()) || Array.from(desiredByFormId.values()).includes(s.getName()));

    // Group response(-ish) sheets by linked form ID, when present.
    const sheetsByFormId = new Map<string, GoogleAppsScript.Spreadsheet.Sheet[]>();
    const unlinked: GoogleAppsScript.Spreadsheet.Sheet[] = [];
    responseishSheets.forEach((sheet) => {
      let formId: string | null = null;
      try {
        formId = extractFormIdFromUrl(sheet.getFormUrl() || '');
      } catch {
        formId = null;
      }
      if (!formId) {
        unlinked.push(sheet);
        return;
      }
      const arr = sheetsByFormId.get(formId) || [];
      arr.push(sheet);
      sheetsByFormId.set(formId, arr);
    });

    // For each known SHAMROCK form, ensure its linked response sheet has the desired name.
    forms.forEach(({ formId, desiredSheetName }) => {
      const linked = sheetsByFormId.get(formId) || [];
      if (linked.length === 0) {
        Log.warn(`No response sheet currently linked to formId=${formId} to rename to '${desiredSheetName}'`);
        return;
      }

      // Prefer a sheet already named correctly.
      const primary = linked.find((s) => s.getName() === desiredSheetName) || linked[0];
      if (primary.getName() !== desiredSheetName) {
        Log.info(`Renaming linked response sheet ${primary.getName()} -> ${desiredSheetName} (formId=${formId})`);
        try {
          primary.setName(desiredSheetName);
        } catch (err) {
          Log.warn(`Unable to rename response sheet to '${desiredSheetName}'. Error: ${err}`);
        }
      }

      // Any other linked sheets for the same form are likely historical destination churn; archive their names.
      linked
        .filter((s) => s.getSheetId() !== primary.getSheetId())
        .forEach((s) => {
          if (/^Archived - /i.test(s.getName())) return;
          const archivedName = `Archived - ${desiredSheetName} (${s.getName()})`;
          try {
            Log.warn(`Archiving extra linked response sheet ${s.getName()} -> ${archivedName} (formId=${formId})`);
            s.setName(archivedName);
          } catch (err) {
            Log.warn(`Unable to archive response sheet ${s.getName()}. Error: ${err}`);
          }
        });
    });

    // For unlinked "Form Responses" sheets, just archive them so they stop looking active.
    unlinked.forEach((s) => {
      if (!/^Form Responses/i.test(s.getName())) return;
      const archivedName = `Archived - ${s.getName()}`;
      try {
        Log.warn(`Archiving unlinked response sheet ${s.getName()} -> ${archivedName}`);
        s.setName(archivedName);
      } catch {
        // Ignore name collisions or protected states.
      }
    });
  }

  function ensureSpreadsheetTrigger(handlerName: string, spreadsheetId: string, event: 'open' | 'edit') {
    if (!spreadsheetId) return;
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some((t) => t.getHandlerFunction() === handlerName && t.getTriggerSourceId?.() === spreadsheetId);
    if (exists) return;
    Log.info(`Creating ${event} trigger for handler=${handlerName} spreadsheet=${spreadsheetId}`);
    const builder = ScriptApp.newTrigger(handlerName).forSpreadsheet(spreadsheetId);
    if (event === 'open') {
      builder.onOpen().create();
    } else {
      builder.onEdit().create();
    }
  }

  function removeDefaultSheetIfPresent(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, allowedNames: Set<string>) {
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet && !allowedNames.has('Sheet1')) {
      // Only remove if there is more than one sheet to avoid deleting the last sheet in a spreadsheet.
      if (spreadsheet.getSheets().length > 1) {
        Log.info(`Removing default sheet 'Sheet1' from spreadsheet=${spreadsheet.getId()}`);
        spreadsheet.deleteSheet(defaultSheet);
      } else {
        Log.warn(`Default sheet 'Sheet1' present but is the only sheet; skipping delete in spreadsheet=${spreadsheet.getId()}`);
      }
    }
  }

  function ensureForm(
    kind: 'attendance' | 'excusal' | 'directory',
    name: string,
    propertyKey: string,
    destinationSpreadsheetId?: string,
  ): Types.EnsureFormResult {
    Log.info(`Ensuring form kind=${kind}`);
    const props = Config.scriptProperties();
    const existingId = props.getProperty(propertyKey);
    let form: GoogleAppsScript.Forms.Form | null = null;
    let created = false;

    if (existingId) {
      try {
        form = FormApp.openById(existingId);
        Log.info(`Found existing form id=${existingId}`);
      } catch (err) {
        Log.warn(`Stored form id invalid for key=${propertyKey}; creating new. Error: ${err}`);
      }
    }

    if (!form) {
      form = FormApp.create(name);
      created = true;
      props.setProperty(propertyKey, form.getId());
      Log.info(`Created form name=${name} id=${form.getId()}`);
    }

    // Keep form title stable (helps ops/debugging).
    try {
      if (form.getTitle() !== name) form.setTitle(name);
    } catch (err) {
      Log.warn(`Unable to set form title. Error: ${err}`);
    }

    // Enforce responder email collection and login requirement (verified identity).
    form.setCollectEmail(true);
    try {
      form.setRequireLogin(true);
    } catch (err) {
      // setRequireLogin is not supported for consumer accounts; log and continue.
      Log.warn(`setRequireLogin not supported in this environment; continuing without it. Error: ${err}`);
    }

    // Response edit policy per form type.
    try {
      if (kind === 'directory') {
        form.setAllowResponseEdits(true);
      } else {
        form.setAllowResponseEdits(false);
      }
    } catch (err) {
      Log.warn(`setAllowResponseEdits not supported in this environment; continuing without it. Error: ${err}`);
    }

    // Route responses into the backend spreadsheet when provided, with retries.
    let destinationConfigured = false;
    if (destinationSpreadsheetId) {
      const currentDestinationId = getFormDestinationSpreadsheetId(form);
      if (currentDestinationId && currentDestinationId === destinationSpreadsheetId) {
        destinationConfigured = true;
      } else {
        if (currentDestinationId && currentDestinationId !== destinationSpreadsheetId) {
          Log.warn(
            `Form destination differs from desired; updating. current=${currentDestinationId} desired=${destinationSpreadsheetId} formId=${form.getId()}`,
          );
        }

        const attemptSetDestination = () => {
        try {
          form.setDestination(FormApp.DestinationType.SPREADSHEET, destinationSpreadsheetId);
          return true;
        } catch (err) {
          Log.warn(`Unable to set form destination to spreadsheet=${destinationSpreadsheetId}. Error: ${err}`);
          return false;
        }
      };

        // Only set destination when needed; retry to handle transient failures.
        for (let i = 0; i < 3 && !destinationConfigured; i++) {
          if (attemptSetDestination()) {
            destinationConfigured = true;
            break;
          }
          Utilities.sleep(500);
        }
      }
    }

    // Note: Built-in "email a copy of responses" setting is not reliably controllable via Apps Script.
    // Future: implement onFormSubmit email receipt as part of the form handler.
    if (kind === 'directory') {
      try {
        form.setConfirmationMessage(
          'Thanks! Please save your response edit link from the confirmation screen so you can update your information later.',
        );
      } catch (err) {
        Log.warn(`Unable to set confirmation message. Error: ${err}`);
      }
    }

    // Seed questions if empty.
    if (kind === 'attendance') FormService.ensureAttendanceForm(form);
    if (kind === 'excusal') FormService.ensureExcusalForm(form);
    if (kind === 'directory') FormService.ensureDirectoryForm(form);

    // Ensure response sheet is sensibly named when destination is set.
    if (destinationSpreadsheetId && destinationConfigured) {
      const desired =
        kind === 'attendance'
          ? Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET
          : kind === 'excusal'
          ? Config.RESOURCE_NAMES.EXCUSAL_FORM_SHEET
          : Config.RESOURCE_NAMES.DIRECTORY_FORM_SHEET;
      ensureResponseSheetNameWithRetry(destinationSpreadsheetId, desired, 10, 1000);
    } else if (destinationSpreadsheetId && !destinationConfigured) {
      Log.warn(`Response destination not configured for form id=${form.getId()}; cannot rename response sheet.`);
    }

    return {
      kind,
      id: form.getId(),
      created,
      url: form.getEditUrl(),
    };
  }

  export function applyFrontendFormatting() {
    const { frontendId } = getIds();
    if (!frontendId) return;
    FrontendFormattingService.applyAll(frontendId);
  }

  export function reapplyFrontendProtections() {
    const { frontendId } = getIds();
    if (!frontendId) return;
    ProtectionService.applyFrontendProtections(frontendId);
  }

  export function toggleFrontendFormatting() {
    const props = Config.scriptProperties();
    const current = String(props.getProperty('DISABLE_FRONTEND_FORMATTING') || '').toLowerCase();
    const next = current === 'true' ? '' : 'true';
    props.setProperty('DISABLE_FRONTEND_FORMATTING', next);
    const status = next === 'true' ? 'OFF (disabled)' : 'ON (enabled)';
    const msg = `Frontend formatting is now ${status}.`;
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (err) {
      Log.info(msg);
    }
  }

  export function refreshDataLegendAndFrontend() {
    DataLegendService.refreshLegendFromArrays();
    SyncService.syncByBackendSheetName('Data Legend');
    applyFrontendFormatting();
  }

  export function applyAttendanceBackendFormattingPublic() {
    applyAttendanceBackendFormatting();
  }

  export function syncDirectoryFrontend() {
    const { frontendId } = getIds();
    if (frontendId) DirectoryService.protectFrontendDirectory(frontendId);
    DirectoryService.syncDirectoryFrontend();
  }

  export function rebuildAttendanceMatrix() {
    AttendanceService.rebuildMatrix();
  }

  export function refreshAttendanceForm() {
    const { backendId } = getIds();
    ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();
  }

  export function rebuildAttendanceForm() {
    const { backendId } = getIds();
    const ensured = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.rebuildAttendanceForm(form);
    // After rebuilding questions, refresh event list and clean up response artifacts.
    FormService.refreshAttendanceFormEventChoices(form);
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();
  }

  export function refreshAttendanceFormEventChoices() {
    const { backendId } = getIds();
    const ensured = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backendId);
    const form = FormApp.openById(ensured.id);
    FormService.refreshAttendanceFormEventChoices(form);
  }

  export function pruneAttendanceResponseColumns() {
    pruneAttendanceResponseColumnsExplicit();
  }

  export function refreshEventsArtifacts() {
    SyncService.syncByBackendSheetName('Events Backend');
    rebuildAttendanceMatrix();
    refreshAttendanceFormEventChoices();
    applyFrontendFormatting();
  }

  export function archiveCoreSheets() {
    const { frontendId, backendId } = getIds();
    const frontendNames = ['Leadership', 'Directory', 'Attendance'];
    const backendNames = ['Leadership Backend', 'Directory Backend', 'Attendance Backend'];

    archiveAndResetSheets(frontendId, Schemas.FRONTEND_TABS, frontendNames);
    archiveAndResetSheets(backendId, Schemas.BACKEND_TABS, backendNames);

    if (frontendId) {
      ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
        ensureTableForSheet(frontendId, name, name.replace(/\s+/g, '_').toLowerCase());
      });
      FrontendFormattingService.applyAll(frontendId);
      ProtectionService.applyFrontendProtections(frontendId);
    }

    if (backendId) {
      applyAttendanceBackendFormatting();
    }
  }

  export function restoreCoreSheetsFromArchive() {
    const { frontendId, backendId } = getIds();
    const frontendNames = ['Leadership', 'Directory', 'Attendance'];
    const backendNames = ['Leadership Backend', 'Directory Backend', 'Attendance Backend'];

    restoreFromArchiveSheets(frontendId, Schemas.FRONTEND_TABS, frontendNames);
    restoreFromArchiveSheets(backendId, Schemas.BACKEND_TABS, backendNames);

    if (frontendId) {
      ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
        ensureTableForSheet(frontendId, name, name.replace(/\s+/g, '_').toLowerCase());
      });
      FrontendFormattingService.applyAll(frontendId);
      ProtectionService.applyFrontendProtections(frontendId);
    }

    if (backendId) {
      applyAttendanceBackendFormatting();
    }
  }

  export function runSetup(): Types.SetupSummary {
    Log.info('Starting setup (ensure-exists)');
    const spreadsheetResults: Types.EnsureSpreadsheetResult[] = [];
    const sheetResults: Types.EnsureSheetResult[] = [];
    const formResults: Types.EnsureFormResult[] = [];

    // Ensure spreadsheets.
    const frontend = ensureSpreadsheet('frontend', Config.RESOURCE_NAMES.FRONTEND_SPREADSHEET, Config.PROPERTY_KEYS.FRONTEND_SHEET_ID);
    const backend = ensureSpreadsheet('backend', Config.RESOURCE_NAMES.BACKEND_SPREADSHEET, Config.PROPERTY_KEYS.BACKEND_SHEET_ID);
    spreadsheetResults.push(frontend, backend);

    // Ensure frontend sheets.
    const frontendSheet = SpreadsheetApp.openById(frontend.id);
    Schemas.FRONTEND_TABS.forEach((tab) => {
      sheetResults.push(ensureSheet(frontendSheet, tab));
    });
    removeDefaultSheetIfPresent(frontendSheet, new Set(Schemas.FRONTEND_TABS.map((t) => t.name)));

    // Ensure backend sheets.
    const backendSheet = SpreadsheetApp.openById(backend.id);
    Schemas.BACKEND_TABS.forEach((tab) => {
      sheetResults.push(ensureSheet(backendSheet, tab));
    });
    removeDefaultSheetIfPresent(backendSheet, new Set(Schemas.BACKEND_TABS.map((t) => t.name)));

    // Ensure forms.
    const attendanceForm = ensureForm('attendance', Config.RESOURCE_NAMES.ATTENDANCE_FORM, Config.PROPERTY_KEYS.ATTENDANCE_FORM_ID, backend.id);
    const excusalForm = ensureForm('excusal', Config.RESOURCE_NAMES.EXCUSAL_FORM, Config.PROPERTY_KEYS.EXCUSAL_FORM_ID, backend.id);
    const directoryForm = ensureForm('directory', Config.RESOURCE_NAMES.DIRECTORY_FORM, Config.PROPERTY_KEYS.DIRECTORY_FORM_ID, backend.id);
    formResults.push(attendanceForm, excusalForm, directoryForm);

    // Normalize response sheet names based on the form actually linked to each sheet.
    normalizeResponseSheetsForForms(backend.id, [
      { formId: attendanceForm.id, desiredSheetName: Config.RESOURCE_NAMES.ATTENDANCE_FORM_SHEET },
      { formId: excusalForm.id, desiredSheetName: Config.RESOURCE_NAMES.EXCUSAL_FORM_SHEET },
      { formId: directoryForm.id, desiredSheetName: Config.RESOURCE_NAMES.DIRECTORY_FORM_SHEET },
    ]);
    // Slim attendance response sheet to drop stale/duplicate columns left over from form rebuilds (keeps any columns that still have data).
    slimAttendanceResponseSheet();
    pruneAttendanceResponseColumnsExplicit();
    normalizeAttendanceBackendHeaders();
    applyAttendanceBackendFormatting();

    // Ensure form submit triggers for receipts/processing.
    ensureFormTrigger('onAttendanceFormSubmit', attendanceForm.id);
    ensureFormTrigger('onExcusalFormSubmit', excusalForm.id);
    ensureFormTrigger('onDirectoryFormSubmit', directoryForm.id);

    // Refresh Data Legend from canonical arrays and sync to frontend.
    refreshDataLegendAndFrontend();

    // Protect user-facing directory and sync it from backend.
    ProtectionService.applyFrontendProtections(frontend.id);
    DirectoryService.syncDirectoryFrontend();

    // Apply frontend validations/banding after syncs.
    FrontendFormattingService.applyAll(frontend.id);

    // Create structured tables on key frontend sheets via Sheets API.
    ['Directory', 'Leadership', 'Attendance', 'Data Legend'].forEach((name) => {
      ensureTableForSheet(frontend.id, name, name.replace(/\s+/g, '_').toLowerCase());
    });

    // Build attendance matrix initially.
    rebuildAttendanceMatrix();

    // Install onOpen triggers for menus and onEdit trigger for backend directory sync.
    ensureSpreadsheetTrigger('onFrontendOpen', frontend.id, 'open');
    ensureSpreadsheetTrigger('onBackendOpen', backend.id, 'open');
    ensureSpreadsheetTrigger('onBackendEdit', backend.id, 'edit');

    Log.info(`Setup finished: spreadsheets=${spreadsheetResults.length}, sheets=${sheetResults.length}, forms=${formResults.length}`);

    return {
      spreadsheets: spreadsheetResults,
      sheets: sheetResults,
      forms: formResults,
    };
  }
}
