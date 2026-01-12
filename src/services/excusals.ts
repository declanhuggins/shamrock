// Excusals processing service: handle notifications, decisions, and management panel.

namespace ExcusalsService {
  /**
   * Send notification email to squadron commander when new excusal submitted.
   */
  export function notifySquadronCommanderOfNewExcusal(excusalRow: Record<string, any>) {
    const squadron = String(excusalRow['squadron'] || '').trim();
    if (!squadron) {
      Log.warn('Cannot notify: excusal has no squadron');
      return;
    }

    const commanderEmail = getSquadronCommanderEmail(squadron);
    if (!commanderEmail) {
      Log.warn(`Cannot notify: no squadron commander email found for ${squadron}`);
      return;
    }

    const lastName = String(excusalRow['last_name'] || '');
    const firstName = String(excusalRow['first_name'] || '');
    const cadetEmail = String(excusalRow['email'] || '');
    const event = String(excusalRow['event'] || '');
    const reason = String(excusalRow['notes'] || '');
    const submittedAt = excusalRow['submitted_at'] ? new Date(excusalRow['submitted_at']) : new Date();

    // Determine time of day
    const hours = submittedAt.getHours();
    const timeOfDay = hours < 12 ? 'Good morning' : hours < 18 ? 'Good afternoon' : 'Good evening';

    // Get commander name
    const commander = lookupLeadershipByEmail(commanderEmail);
    const commanderLastName = String(commander?.last_name || 'Commander').trim();

    const managementSheetUrl = getManagementSpreadsheetUrl();

    const subject = `New Excusal Request Submitted: ${lastName}, ${firstName} – ${event}`;
    const body = `${timeOfDay} C/${commanderLastName},

You have received a new excusal request from Cadet ${firstName} ${lastName}.

Details:
• Cadet: ${lastName}, ${firstName} (${cadetEmail})
• Event: ${event}
• Reason: ${reason}

Review & take action here:
${managementSheetUrl}

Very respectfully,
SHAMROCK Automations`;

    try {
      GmailApp.sendEmail(commanderEmail, subject, body, {
        name: 'SHAMROCK Automations',
        replyTo: cadetEmail,
        cc: cadetEmail,
      });
      Log.info(`Excusal notification sent to ${commanderEmail} for ${lastName}, ${firstName}`);
    } catch (err) {
      Log.warn(`Failed to send excusal notification to ${commanderEmail}: ${err}`);
    }
  }

  /**
   * Update attendance matrix when excusal is submitted.
   * Empty cell -> ER (Excusal Requested)
   * Unexcused (U) -> UR (Unexcused Report Submitted)
   */
  export function updateAttendanceOnExcusalSubmission(excusalRow: Record<string, any>) {
    const backendId = Config.getBackendId();
    if (!backendId) return;

    const lastName = String(excusalRow['last_name'] || '').trim();
    const firstName = String(excusalRow['first_name'] || '').trim();
    const eventName = String(excusalRow['event'] || '').trim();

    if (!lastName || !firstName || !eventName) return;

    // Determine current matrix value to pick ER vs UR
    const current = lookupMatrixValue(eventName, lastName, firstName);
    const code = current === 'U' ? 'UR' : 'ER';

    const logEntry = {
      submission_id: `excusal-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      submitted_at: new Date(),
      event: eventName,
      attendance_type: code,
      email: excusalRow['email'] || '',
      name: 'Excusal Request',
      flight: excusalRow['flight'] || '',
      cadets: `${lastName}, ${firstName}`,
    };

    appendAttendanceLogs([logEntry]);
    AttendanceService.applyAttendanceLogEntry(logEntry);
  }

  /**
   * Get squadron commander email by squadron name.
   */
  function getSquadronCommanderEmail(squadron: string): string {
    const backendId = Config.getBackendId();
    if (!backendId) return '';

    const leadershipSheet = SheetUtils.getSheet(backendId, 'Leadership Backend');
    if (!leadershipSheet) return '';

    const table = SheetUtils.readTable(leadershipSheet);
    const squadronNormalized = squadron.toLowerCase().trim();

    const commander = table.rows.find((row) => {
      const role = String(row['role'] || '').toLowerCase().trim();
      const rowSquadron = String(row['squadron'] || '').toLowerCase().trim();
      const roleIncludesSquadron = Arrays.SQUADRONS.some((sq) => {
        const sqLower = sq.toLowerCase();
        return role.includes(sqLower) && squadronNormalized === sqLower;
      });
      const matchesSquadron = rowSquadron ? rowSquadron === squadronNormalized : roleIncludesSquadron;
      return role.includes('squadron commander') && matchesSquadron;
    });

    return commander ? String(commander['email'] || '') : '';
  }

  /**
   * Look up leadership entry by email.
   */
  function lookupLeadershipByEmail(email: string): Record<string, any> | null {
    const backendId = Config.getBackendId();
    if (!backendId || !email) return null;

    const leadershipSheet = SheetUtils.getSheet(backendId, 'Leadership Backend');
    if (!leadershipSheet) return null;

    const table = SheetUtils.readTable(leadershipSheet);
    const lower = email.toLowerCase();
    return table.rows.find((r) => String(r['email'] || '').toLowerCase() === lower) || null;
  }

  /**
   * Get or create the excusals management spreadsheet.
   */
  export function ensureManagementSpreadsheet(): string {
    const props = Config.scriptProperties();
    const existingId = props.getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');

    if (existingId) {
      try {
        const ss = SpreadsheetApp.openById(existingId);
        // Ensure squadron sheets exist and are initialized even if spreadsheet already exists.
        const squadrons = Arrays.SQUADRONS.filter((s) => s !== 'Abroad');
        squadrons.forEach((squadron) => {
          let sheet = ss.getSheetByName(squadron);
          if (!sheet) {
            sheet = ss.insertSheet(squadron);
          }
          initializeSquadronManagementSheet(sheet, squadron);
        });
        return existingId;
      } catch (err) {
        Log.warn(`Stored management spreadsheet ID invalid; creating new. Error: ${err}`);
      }
    }

    // Create new management spreadsheet
    const ss = SpreadsheetApp.create('SHAMROCK Excusals Management');
    const newId = ss.getId();
    props.setProperty('EXCUSALS_MANAGEMENT_SHEET_ID', newId);
    Log.info(`Created excusals management spreadsheet: ${newId}`);

    // Create squadron sheets first (before deleting default)
    // Use squadrons from canonical Arrays, excluding 'Abroad'
    const squadrons = Arrays.SQUADRONS.filter((s) => s !== 'Abroad');
    squadrons.forEach((squadron) => {
      const sheet = ss.insertSheet(squadron);
      initializeSquadronManagementSheet(sheet, squadron);
    });

    // Remove default sheet only after new sheets exist
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);

    return newId;
  }

  /**
   * Initialize a squadron management sheet with headers and structure.
   */
  function initializeSquadronManagementSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet, squadron: string) {
    const schema = Schemas.EXCUSALS_MANAGEMENT_SCHEMA;
    const machineHeaders = schema.machineHeaders!;
    const displayHeaders = schema.displayHeaders!;
    const headerWidth = machineHeaders.length;

    // Set machine headers in row 1
    sheet.getRange(1, 1, 1, headerWidth).setValues([machineHeaders]);
    sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setHorizontalAlignment('center');

    // Set display headers in row 2
    sheet.getRange(2, 1, 1, headerWidth).setValues([displayHeaders]);
    sheet.getRange(2, 1, 1, headerWidth).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e8e8e8');

    // Hide machine headers (row 1)
    sheet.hideRows(1, 1);
    // Ensure minimal rows = 3 (2 headers + 1 blank data row)
    const lastRow = sheet.getLastRow();
    const maxRows = sheet.getMaxRows();
    const MIN_ROWS = 3;
    if (lastRow <= 2) {
      if (maxRows < MIN_ROWS) {
        sheet.insertRowsAfter(maxRows, MIN_ROWS - maxRows);
      } else if (maxRows > MIN_ROWS) {
        sheet.deleteRows(MIN_ROWS + 1, maxRows - MIN_ROWS);
      }
    }

    // Trim any extra columns beyond the schema width to keep things tidy.
    const maxCols = sheet.getMaxColumns();
    if (maxCols > headerWidth) {
      sheet.deleteColumns(headerWidth + 1, maxCols - headerWidth);
    }

    // Set column widths
    sheet.setColumnWidth(1, 150); // timestamp
    sheet.setColumnWidth(2, 100); // decision
    sheet.setColumnWidth(3, 150); // event
    sheet.setColumnWidth(4, 220); // reason
    sheet.setColumnWidth(5, 150); // email
    sheet.setColumnWidth(6, 100); // last_name
    sheet.setColumnWidth(7, 100); // first_name
    sheet.setColumnWidth(8, 80);  // flight
    sheet.setColumnWidth(9, 120); // request_id

    // Freeze first two rows (machine headers + display headers)
    sheet.setFrozenRows(2);

    // Add data validation for Decision column (col 2, starting at row 3 since row 2 is frozen)
    const decisionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Approved', 'Denied'])
      .setHelpText('Select Approved or Denied')
      .build();
    sheet.getRange('B3:B').setDataValidation(decisionRule);

    Log.info(`Initialized management sheet for ${squadron} squadron`);
  }

  /**
   * Sync excusal to management spreadsheet.
   */
  export function syncExcusalToManagementPanel(excusalRow: Record<string, any>) {
    const squadron = String(excusalRow['squadron'] || '').trim();
    if (!squadron) {
      Log.warn('Cannot sync excusal to management panel: no squadron');
      return;
    }

    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');
    if (!managementId) {
      Log.warn('Excusals management spreadsheet not found; skipping sync');
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(managementId);
      const sheet = ss.getSheetByName(squadron);
      if (!sheet) {
        Log.warn(`Sheet for squadron ${squadron} not found in management spreadsheet`);
        return;
      }

      // Ensure sheet has capacity and append after existing data (starting at row 3)
      const nextRow = Math.max(3, sheet.getLastRow() + 1);
      const maxRows = sheet.getMaxRows();
      if (nextRow > maxRows) {
        sheet.insertRowsAfter(maxRows, nextRow - maxRows);
      }

      // Ensure columns match current schema (handles legacy sheets created before Reason column was added)
      const maxCols = sheet.getMaxColumns();
      const requiredCols = Schemas.EXCUSALS_MANAGEMENT_SCHEMA.machineHeaders!.length;
      if (maxCols < requiredCols) {
        sheet.insertColumnsAfter(maxCols, requiredCols - maxCols);
      }

      // Refresh headers to match schema (machine headers row 1, display headers row 2)
      const machineHeaders = Schemas.EXCUSALS_MANAGEMENT_SCHEMA.machineHeaders!;
      const displayHeaders = Schemas.EXCUSALS_MANAGEMENT_SCHEMA.displayHeaders!;
      sheet.getRange(1, 1, 1, machineHeaders.length).setValues([machineHeaders]);
      sheet.getRange(2, 1, 1, displayHeaders.length).setValues([displayHeaders]);

      const rowData = [
        excusalRow['submitted_at'] || '',
        '', // Decision column starts empty
        excusalRow['event'] || '',
        excusalRow['reason'] || excusalRow['notes'] || '',
        excusalRow['email'] || '',
        excusalRow['last_name'] || '',
        excusalRow['first_name'] || '',
        excusalRow['flight'] || '',
        excusalRow['request_id'] || '',
      ];

      sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

      // Trim extra rows and columns, then sort by Timestamp (descending)
      trimAndSortManagementSheet(sheet);

      // Reapply protections so new rows remain covered
      applyManagementSheetProtections(ss);

      Log.info(`Synced excusal ${excusalRow['request_id']} to ${squadron} management sheet`);
    } catch (err) {
      Log.warn(`Failed to sync excusal to management panel: ${err}`);
    }
  }

  /**
   * Trim empty rows and columns from management sheet, sort by Timestamp descending.
   */
  function trimAndSortManagementSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    // Only sort and trim if there's any rows present
    if (lastRow >= 2) {
      const dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
      const values = dataRange.getValues();

      // Preserve first two rows (machine + display headers)
      const headerRows = values.slice(0, 2);
      const dataRows = values.slice(2);

      // Sort data rows by timestamp descending (col 1)
      dataRows.sort((a: any[], b: any[]) => {
        const timeA = new Date(a[0] || '').getTime();
        const timeB = new Date(b[0] || '').getTime();
        return timeB - timeA; // Descending (latest first)
      });

      const sortedValues = [...headerRows, ...dataRows];
      dataRange.setValues(sortedValues);

      // Delete extra rows beyond data, but keep minimum of 3 total rows
      const MIN_ROWS = 3;
      const targetRows = Math.max(lastRow, MIN_ROWS);
      if (maxRows > targetRows) {
        sheet.deleteRows(targetRows + 1, maxRows - targetRows);
      } else if (maxRows < targetRows) {
        sheet.insertRowsAfter(maxRows, targetRows - maxRows);
      }
    }

    // Delete extra columns beyond data
    if (lastColumn < maxCols) {
      sheet.deleteColumns(lastColumn + 1, maxCols - lastColumn);
    }
  }

  /**
   * Get the management spreadsheet URL.
   */
  function getManagementSpreadsheetUrl(): string {
    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');
    if (!managementId) return '(Management panel URL unavailable)';
    return `https://docs.google.com/spreadsheets/d/${managementId}`;
  }

  /**
   * Share management spreadsheet with squadron and flight commanders, apply sheet protections.
   */
  export function shareAndProtectManagementSpreadsheet() {
    const managementId = Config.scriptProperties().getProperty('EXCUSALS_MANAGEMENT_SHEET_ID');
    if (!managementId) {
      Log.warn('Excusals management spreadsheet not found; cannot share or protect');
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(managementId);
      const backendId = Config.getBackendId();
      if (!backendId) {
        Log.warn('Cannot share management spreadsheet: backend ID missing');
        return;
      }

      // Get squadron and flight commander emails from Leadership Backend
      const leadershipSheet = SheetUtils.getSheet(backendId, 'Leadership Backend');
      if (!leadershipSheet) {
        Log.warn('Cannot share management spreadsheet: Leadership Backend not found');
        return;
      }

      const table = SheetUtils.readTable(leadershipSheet);
      const commanderEmails = new Set<string>();
      table.rows.forEach((row) => {
        const role = String(row['role'] || '').toLowerCase().trim();
        const email = String(row['email'] || '').trim();
        if ((role.includes('squadron commander') || role.includes('flight commander')) && email) {
          commanderEmails.add(email);
        }
      });

      if (commanderEmails.size === 0) {
        Log.warn('No squadron or flight commanders found in Leadership Backend');
        return;
      }

      // Share spreadsheet with all commanders
      const editorsArray = Array.from(commanderEmails);
      editorsArray.forEach((email) => {
        try {
          ss.addEditor(email);
        } catch (err) {
          Log.warn(`Failed to add editor ${email}: ${err}`);
        }
      });

      // Apply range protections: each squadron sheet editable only by its commander
      applyManagementSheetProtections(ss);

      Log.info(`Shared management spreadsheet with ${editorsArray.length} commanders and applied protections`);
    } catch (err) {
      Log.warn(`Failed to share and protect management spreadsheet: ${err}`);
    }
  }

  /**
   * Handle edits to the Excusals Backend (e.g., decision approval/denial).
   */
  export function handleExcusalsBackendEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    const range = e?.range;
    if (!range) return;

    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();
    const newValue = String((e as any)?.value ?? range.getValue() ?? '').trim();

    // Get headers to find Decision column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim().toLowerCase());
    const decisionColIdx = headers.indexOf('decision');
    const statusIdx = headers.indexOf('status');
    const decidedByIdx = headers.indexOf('decided_by');
    const decidedAtIdx = headers.indexOf('decided_at');

    // Only process if Decision column was edited and value is Approved/Denied
    if (col - 1 !== decisionColIdx || row < 3) return;
    if (!['Approved', 'Denied'].includes(newValue)) return;

    try {
      const backendSheet = SheetUtils.getSheet(Config.getBackendId(), 'Excusals Backend');
      if (!backendSheet) return;
      const table = SheetUtils.readTable(backendSheet);
      const rowData = table.rows[row - 2]; // -1 for header, -1 for 0-index
      if (!rowData) return;

      const oldDecision = String(rowData['decision'] || '').trim();
      const cadetEmail = String(rowData['email'] || '').trim();
      const eventName = String(rowData['event'] || '').trim();
      const reason = String(rowData['notes'] || '').trim();
      const squadron = String(rowData['squadron'] || '').trim();
      const firstName = String(rowData['first_name'] || '').trim();
      const lastName = String(rowData['last_name'] || '').trim();

      // Update status and decided_at in the same row
      if (statusIdx >= 0) sheet.getRange(row, statusIdx + 1).setValue(newValue === 'Approved' ? 'Approved' : 'Denied');
      if (decidedByIdx >= 0) sheet.getRange(row, decidedByIdx + 1).setValue(Session.getActiveUser().getEmail());
      if (decidedAtIdx >= 0) sheet.getRange(row, decidedAtIdx + 1).setValue(new Date().toISOString());

      // Update attendance matrix based on decision
      updateAttendanceOnExcusalDecision({
        lastName,
        firstName,
        eventName,
        decision: newValue,
      });

      // Check if decision is being changed (not initial decision)
      const isDecisionChange = oldDecision && oldDecision !== newValue;

      // Send decision email to cadet from squadron commander
      sendExcusalDecisionEmail({
        cadetEmail,
        cadetFirstName: firstName,
        cadetLastName: lastName,
        event: eventName,
        decision: newValue,
        previousDecision: isDecisionChange ? oldDecision : undefined,
        reason,
        squadron,
      });

      Log.info(`Excusal decision recorded: row ${row} -> ${newValue}${isDecisionChange ? ` (changed from ${oldDecision})` : ''}`);
    } catch (err) {
      Log.warn(`Failed to handle Excusals Backend edit: ${err}`);
    }
  }

  /**
   * Update attendance matrix when excusal decision is made.
   * Approved: ER->E, UR->E
   * Denied: ER->ED, UR->U
   */
  function updateAttendanceOnExcusalDecision(opts: {
    lastName: string;
    firstName: string;
    eventName: string;
    decision: string;
  }) {
    const current = lookupMatrixValue(opts.eventName, opts.lastName, opts.firstName);
    let code = '';
    if (opts.decision === 'Approved') {
      code = 'E';
    } else {
      code = current === 'UR' ? 'U' : 'ED';
    }

    const logEntry = {
      submission_id: `excusal-decision-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      submitted_at: new Date(),
      event: opts.eventName,
      attendance_type: code,
      email: 'excusal-decision',
      name: 'Excusal Decision',
      flight: '',
      cadets: `${opts.lastName}, ${opts.firstName}`,
    };

    appendAttendanceLogs([logEntry]);
    AttendanceService.applyAttendanceLogEntry(logEntry);
  }

  /**
   * Send decision notification email to cadet from squadron commander.
   */
  function sendExcusalDecisionEmail(opts: {
    cadetEmail: string;
    cadetFirstName: string;
    cadetLastName: string;
    event: string;
    decision: string;
    previousDecision?: string;
    reason: string;
    squadron: string;
  }) {
    if (!opts.cadetEmail) {
      Log.warn('Cannot send decision email: no cadet email');
      return;
    }

    const backendId = Config.getBackendId();
    if (!backendId) {
      Log.warn('Cannot send decision email: no backend ID');
      return;
    }

    // Get squadron commander details
    const commanderEmail = getSquadronCommanderEmail(opts.squadron);
    const commander = commanderEmail ? lookupLeadershipByEmail(commanderEmail) : null;
    const commanderLastName = String(commander?.last_name || 'Commander').trim();

    // Determine time of day
    const hours = new Date().getHours();
    const timeOfDay = hours < 12 ? 'Good morning' : hours < 18 ? 'Good afternoon' : 'Good evening';

    let subject: string;
    let body: string;

    if (opts.previousDecision) {
      // Decision change notification
      subject = `Excusal Request Decision Changed: ${opts.cadetLastName}, ${opts.cadetFirstName} – ${opts.event}`;
      body = `${timeOfDay} Cadet ${opts.cadetFirstName} ${opts.cadetLastName},

Your excusal request for ${opts.event} has been reconsidered and the decision has changed.

Previous decision: ${opts.previousDecision}
New decision: ${opts.decision}

Your excusal reason: ${opts.reason}

If you have questions or would like to appeal, contact your flight/squadron commander through the chain of command.

Very respectfully,
C/${commanderLastName}`;
    } else {
      // Initial decision notification
      subject = `Excusal Request ${opts.decision}: ${opts.cadetLastName}, ${opts.cadetFirstName} – ${opts.event}`;
      body = `${timeOfDay} Cadet ${opts.cadetFirstName} ${opts.cadetLastName},

Your excusal request for ${opts.event} has been ${opts.decision.toLowerCase()}.

Your excusal reason: ${opts.reason}

If you have questions or would like to appeal, contact your flight/squadron commander through the chain of command.

Very respectfully,
C/${commanderLastName}`;
    }

    try {
      GmailApp.sendEmail(opts.cadetEmail, subject, body, {
        name: 'SHAMROCK Automations',
        replyTo: commanderEmail || 'shamrock@nd.edu',
        cc: commanderEmail || undefined,
      });
      Log.info(`Decision email sent to ${opts.cadetEmail}; decision=${opts.decision}${opts.previousDecision ? ` (changed from ${opts.previousDecision})` : ''}`);
    } catch (err) {
      Log.warn(`Failed to send decision email to ${opts.cadetEmail}: ${err}`);
    }
  }

  function appendAttendanceLogs(logs: Record<string, any>[]) {
    if (!logs.length) return;
    const backendId = Config.getBackendId();
    if (!backendId) return;
    const sheet = SheetUtils.getSheet(backendId, 'Attendance Backend');
    if (!sheet) return;
    SheetUtils.appendRows(sheet, logs);
  }

  // Helper: lookup current matrix value for a cadet/event (backend matrix)
  function lookupMatrixValue(eventName: string, lastName: string, firstName: string): string {
    const backendId = Config.getBackendId();
    if (!backendId) return '';
    const matrixSheet = SheetUtils.getSheet(backendId, 'Attendance Matrix Backend');
    if (!matrixSheet) return '';

    const lastRow = matrixSheet.getLastRow();
    const lastCol = matrixSheet.getLastColumn();
    if (lastRow < 3 || lastCol < 1) return '';

    const headers = matrixSheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map((h) => String(h || '').trim());
    const eventColIdx = headers.indexOf(eventName);
    const lastIdx = headers.indexOf('last_name');
    const firstIdx = headers.indexOf('first_name');
    if (eventColIdx < 0 || lastIdx < 0 || firstIdx < 0) return '';

    const data = matrixSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    for (let i = 0; i < data.length; i++) {
      if (
        String(data[i][lastIdx] || '').trim().toLowerCase() === lastName.toLowerCase() &&
        String(data[i][firstIdx] || '').trim().toLowerCase() === firstName.toLowerCase()
      ) {
        return String(data[i][eventColIdx] || '').trim();
      }
    }
    return '';
  }

  /**
   * Apply sheet protections so commanders can only edit their own squadron sheet.
   */
  function applyManagementSheetProtections(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheets = ss.getSheets();
    const squadrons = Arrays.SQUADRONS.filter((s) => s !== 'Abroad');

    sheets.forEach((sheet) => {
      const sheetName = sheet.getName();
      if (!squadrons.includes(sheetName)) return;

      const commanderEmail = getSquadronCommanderEmail(sheetName);
      const lastCol = Math.max(1, sheet.getLastColumn(), sheet.getMaxColumns());
      const maxRows = sheet.getMaxRows();

      // Remove existing protections (sheet and ranges)
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => {
        try { p.remove(); } catch {}
      });
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach((p) => {
        try { p.remove(); } catch {}
      });

      // Protect header rows (1-2): no editors
      const headerRange = sheet.getRange(1, 1, 2, lastCol);
      try {
        const headerProt = headerRange.protect().setDescription(`${sheetName}: Headers protected`);
        headerProt.setWarningOnly(false);
        try { headerProt.removeEditors(headerProt.getEditors()); } catch {}
      } catch (err) {
        Log.warn(`Failed to protect headers on ${sheetName}: ${err}`);
      }

      // Protect data rows (from row 3 down): only squadron commander may edit
      const dataRowCount = Math.max(1, maxRows - 2);
      if (dataRowCount > 0) {
        const dataRange = sheet.getRange(3, 1, dataRowCount, lastCol);
        try {
          const dataProt = dataRange.protect().setDescription(`${sheetName}: Data editable only by squadron commander`);
          dataProt.setWarningOnly(false);
          try { dataProt.removeEditors(dataProt.getEditors()); } catch {}
          if (commanderEmail) {
            try { dataProt.addEditor(commanderEmail); } catch (e) {
              Log.warn(`Failed adding commander editor ${commanderEmail} on ${sheetName}: ${e}`);
            }
          } else {
            Log.warn(`No commander email found for ${sheetName}; data will be owner-only`);
          }
          // Ensure owner also has edit access
          try { dataProt.addEditor(Session.getActiveUser().getEmail()); } catch {}
        } catch (err) {
          Log.warn(`Failed to protect data range on ${sheetName}: ${err}`);
        }
      }

      Log.info(`Applied range protections on ${sheetName}; commander=${commanderEmail || 'none'}`);
    });
  }
}
