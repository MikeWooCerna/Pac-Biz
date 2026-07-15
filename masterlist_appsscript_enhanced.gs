// ============================================================
//  CONFIG  — set once, do not change anything else here
// ============================================================

const LOG_SPREADSHEET_ID = '18hKmm2SmlWqB23osiV3JTF0aWn86vvZ-YJSC-Rr3JcY';
const LOG_ACCOUNT_NAME   = 'Masterlist';

// Stop before Google's 30-minute hard timeout so Movement rows can finish cleanly.
const MAX_RUNTIME_MS = 25 * 60 * 1000;
const HISTORY_REFRESH_MIN_REMAINING_MS = 4 * 60 * 1000;

// Required columns per sheet — health check will flag any that are missing.
const REQUIRED_MASTER_COLS   = ['Emp Name', 'Department', 'LOB / Account', 'Immediate Supervisor', 'Manager', 'Job Title', 'Employment Status', 'Attrition Date'];
const REQUIRED_MOVEMENT_COLS = ['Employee Name', 'Effective Date', 'New Department', 'New Account', 'New Supervisor', 'New Job Title', 'New Employment Status'];
const REQUIRED_HISTORY_COLS  = ['Emp Name', 'Date Generated', 'Change Type', 'Week'];


// ============================================================
//  MAIN ENTRY POINT  — the only function you trigger
// ============================================================

function runMasterlistProcess() {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('runMasterlistProcess skipped: another run is still active.');
    return;
  }

  let rowsProcessed      = 0;
  let historyRowsWritten = 0;
  let status             = 'ok';
  let notes              = '';

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const master   = ss.getSheetByName('Masterlist');
    const movement = ss.getSheetByName('Movement');
    let   history  = ss.getSheetByName('History');
    if (!history) history = ss.insertSheet('History');

    const health = runHealthCheck(ss, master, movement, history);

    if (health.blocking) {
      status = 'error';
      notes  = 'Blocked by health check: ' + health.errors.slice(0, 3).join(' | ');
      Logger.log('runMasterlistProcess BLOCKED: ' + notes);
      return;
    }

    if (health.warnings.length > 0) {
      status = 'warn';
      notes  = health.warnings.slice(0, 3).join(' | ');
    }

    const today     = new Date();
    const todayText = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
    const weekStart = getWeekStart(todayText);

    setupMovementTrackingColumns(movement);
    setupHistorySheet(master, history);

    // IMPORTANT: Movement processing runs before the heavier daily history snapshot.
    // This prevents past-effective attritions from being stranded when the script is close to timeout.
    const result  = processMovements(master, movement, history, today, todayText, startTime);
    rowsProcessed = result.movementsApplied;

    if (result.warnings.length > 0) {
      status = 'warn';
      notes  = (notes ? notes + ' | ' : '') + result.warnings.slice(0, 3).join(' | ');
    }

    if (result.partial) {
      status = 'warn';
      notes = appendNote(notes, 'Movement processing paused before timeout; next trigger will resume.');
      return;
    }

    if (remainingRuntimeMs(startTime) < HISTORY_REFRESH_MIN_REMAINING_MS) {
      status = 'warn';
      notes = appendNote(notes, 'Skipped daily history snapshot: not enough runtime after movement processing.');
      return;
    }

    historyRowsWritten = refreshTodayHistory(master, history, todayText, weekStart, result.appliedChanges, startTime);
    if (historyRowsWritten === -1) {
      historyRowsWritten = 0;
      status = 'warn';
      notes = appendNote(notes, 'Daily history snapshot paused before timeout; next trigger will retry.');
    }

  } catch (e) {
    status = 'error';
    notes  = e.message;
    Logger.log('runMasterlistProcess ERROR: ' + e.message);
  } finally {
    const durationSec = Math.round((Date.now() - startTime) / 1000);
    logPipelineRun(rowsProcessed, historyRowsWritten, durationSec, status, notes);
    lock.releaseLock();
  }
}


// ============================================================
//  HEALTH CHECK
// ============================================================

function runHealthCheck(ss, master, movement, history) {
  const startTime = Date.now();
  const ts        = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  const errors   = [];
  const warnings = [];
  const checks   = [];
  const stats    = {};

  function addCheck(category, checkName, result, detail) {
    checks.push({ category, checkName, result, detail });
    if (result === 'ERROR')   errors.push(checkName + ': ' + detail);
    if (result === 'WARNING') warnings.push(checkName + ': ' + detail);
  }

  const masterExists   = !!master;
  const movementExists = !!movement;
  const historyExists  = !!history;

  addCheck('Sheets', 'Masterlist sheet exists',  masterExists   ? 'OK'    : 'ERROR',   masterExists   ? 'Found' : 'Sheet not found — create a tab named "Masterlist"');
  addCheck('Sheets', 'Movement sheet exists',    movementExists ? 'OK'    : 'ERROR',   movementExists ? 'Found' : 'Sheet not found — create a tab named "Movement"');
  addCheck('Sheets', 'History sheet exists',     historyExists  ? 'OK'    : 'WARNING', historyExists  ? 'Found' : 'Not found — will be created on first run');

  if (!masterExists || !movementExists) {
    return finalize(ss, ts, checks, errors, warnings, stats, startTime, true);
  }

  const masterHeaders   = getHeaders(master);
  const movementHeaders = getHeaders(movement);
  const historyHeaders  = historyExists ? getHeaders(history) : [];

  const missingMaster = REQUIRED_MASTER_COLS.filter(c => getCol(masterHeaders, c) === -1);
  addCheck('Columns', 'Masterlist required columns',
    missingMaster.length === 0 ? 'OK' : 'ERROR',
    missingMaster.length === 0 ? 'All ' + REQUIRED_MASTER_COLS.length + ' required columns present'
      : 'Missing: ' + missingMaster.join(', ')
  );

  const missingMovement = REQUIRED_MOVEMENT_COLS.filter(c => getCol(movementHeaders, c) === -1);
  addCheck('Columns', 'Movement required columns',
    missingMovement.length === 0 ? 'OK' : 'ERROR',
    missingMovement.length === 0 ? 'All ' + REQUIRED_MOVEMENT_COLS.length + ' required columns present'
      : 'Missing: ' + missingMovement.join(', ')
  );

  if (historyExists && historyHeaders.length > 0) {
    const missingHistory = REQUIRED_HISTORY_COLS.filter(c => getCol(historyHeaders, c) === -1);
    addCheck('Columns', 'History required columns',
      missingHistory.length === 0 ? 'OK' : 'WARNING',
      missingHistory.length === 0 ? 'All ' + REQUIRED_HISTORY_COLS.length + ' required columns present'
        : 'Missing: ' + missingHistory.join(', ')
    );
  }

  const masterLastRow   = master.getLastRow();
  const movementLastRow = movement.getLastRow();
  const historyLastRow  = historyExists ? history.getLastRow() : 0;

  const masterDataRows   = Math.max(0, masterLastRow - 1);
  const movementDataRows = Math.max(0, movementLastRow - 1);
  const historyDataRows  = Math.max(0, historyLastRow - 1);

  stats.masterRows   = masterDataRows;
  stats.movementRows = movementDataRows;
  stats.historyRows  = historyDataRows;

  addCheck('Row Counts', 'Masterlist rows loaded',
    masterDataRows > 0 ? 'OK' : 'ERROR',
    masterDataRows > 0 ? masterDataRows + ' data rows' : 'No data rows found — sheet may be empty'
  );

  addCheck('Row Counts', 'Movement rows loaded',
    movementDataRows > 0 ? 'OK' : 'WARNING',
    movementDataRows > 0 ? movementDataRows + ' data rows' : 'No data rows — nothing to process today'
  );

  addCheck('Row Counts', 'History rows loaded',
    'INFO',
    historyDataRows > 0 ? historyDataRows + ' historical rows on file' : 'Empty — will be populated on first run'
  );

  if (masterDataRows > 0) {
    const masterData         = master.getRange(2, 1, masterDataRows, master.getLastColumn()).getValues();
    const nameCol            = getCol(masterHeaders, 'Emp Name');
    const statusCol          = getCol(masterHeaders, 'Employment Status');
    const supervisorCol      = getCol(masterHeaders, 'Immediate Supervisor');
    const deptCol            = getCol(masterHeaders, 'Department');

    const blankNames = masterData.filter(r => !normalize(r[nameCol])).length;
    addCheck('Masterlist Quality', 'Blank employee names',
      blankNames === 0 ? 'OK' : 'ERROR',
      blankNames === 0 ? 'None found' : blankNames + ' row(s) have no employee name'
    );

    const nameCounts = {};
    masterData.forEach(r => {
      const n = normalize(r[nameCol]);
      if (n) nameCounts[n] = (nameCounts[n] || 0) + 1;
    });
    const dupeNames = Object.entries(nameCounts).filter(([, c]) => c > 1).map(([n]) => n);
    addCheck('Masterlist Quality', 'Duplicate employee names',
      dupeNames.length === 0 ? 'OK' : 'WARNING',
      dupeNames.length === 0 ? 'None found'
        : dupeNames.length + ' duplicate(s): ' + dupeNames.slice(0, 5).join(', ') + (dupeNames.length > 5 ? '…' : '')
    );

    const activeCount   = masterData.filter(r => normalize(r[statusCol]) !== 'INACTIVE').length;
    const inactiveCount = masterData.filter(r => normalize(r[statusCol]) === 'INACTIVE').length;
    stats.activeCount   = activeCount;
    stats.inactiveCount = inactiveCount;
    addCheck('Masterlist Quality', 'Active / inactive headcount',
      'INFO',
      'Active: ' + activeCount + '  |  Inactive: ' + inactiveCount + '  |  Total: ' + masterDataRows
    );

    const blankSupervisors = masterData.filter(r =>
      normalize(r[statusCol]) !== 'INACTIVE' && !String(r[supervisorCol]).trim()
    ).length;
    addCheck('Masterlist Quality', 'Active rows with no supervisor',
      blankSupervisors === 0 ? 'OK' : 'WARNING',
      blankSupervisors === 0 ? 'None found' : blankSupervisors + ' active row(s) missing Immediate Supervisor'
    );

    const blankDepts = masterData.filter(r =>
      normalize(r[statusCol]) !== 'INACTIVE' && !String(r[deptCol]).trim()
    ).length;
    addCheck('Masterlist Quality', 'Active rows with no department',
      blankDepts === 0 ? 'OK' : 'WARNING',
      blankDepts === 0 ? 'None found' : blankDepts + ' active row(s) missing Department'
    );
  }

  if (movementDataRows > 0) {
    const movementData     = movement.getRange(2, 1, movementDataRows, movement.getLastColumn()).getValues();
    const movNameCol       = getCol(movementHeaders, 'Employee Name');
    const movEffDateCol    = getCol(movementHeaders, 'Effective Date');
    const processedCol     = getCol(movementHeaders, 'Processed');
    const voidCol          = getCol(movementHeaders, 'Void');

    const pendingRows = movementData.filter(r => {
      const processed = processedCol === -1 ? '' : normalize(r[processedCol] || '');
      const isVoid    = voidCol === -1 ? '' : normalize(r[voidCol] || '');
      return processed !== 'YES' && isVoid !== 'YES';
    });
    stats.pendingMovements = pendingRows.length;

    addCheck('Movement Quality', 'Pending movements count',
      'INFO',
      pendingRows.length + ' unprocessed, non-voided movement row(s)'
    );

    const today = stripTime(new Date());
    const overdueRows = pendingRows.filter(r => {
      const d = new Date(r[movEffDateCol]);
      return !isNaN(d) && stripTime(d) <= today;
    });
    addCheck('Movement Quality', 'Overdue effective movements',
      overdueRows.length === 0 ? 'OK' : 'WARNING',
      overdueRows.length === 0 ? 'None found' : overdueRows.length + ' pending row(s) have Effective Date today or earlier and should process this run'
    );

    const invalidDates = pendingRows.filter(r => isNaN(new Date(r[movEffDateCol]))).length;
    addCheck('Movement Quality', 'Invalid effective dates',
      invalidDates === 0 ? 'OK' : 'WARNING',
      invalidDates === 0 ? 'None found' : invalidDates + ' pending row(s) have unparseable Effective Date'
    );

    const blankMovNames = pendingRows.filter(r => !normalize(r[movNameCol])).length;
    addCheck('Movement Quality', 'Blank employee names in Movement',
      blankMovNames === 0 ? 'OK' : 'WARNING',
      blankMovNames === 0 ? 'None found' : blankMovNames + ' pending row(s) have no Employee Name'
    );

    if (masterDataRows > 0) {
      const masterData  = master.getRange(2, 1, masterDataRows, master.getLastColumn()).getValues();
      const masterNames = new Set(masterData.map(r => normalize(r[getCol(masterHeaders, 'Emp Name')])));
      const notInMaster = pendingRows.filter(r => {
        const n = normalize(r[movNameCol]);
        return n && !masterNames.has(n);
      }).map(r => String(r[movNameCol]).trim());
      const uniqueNotInMaster = [...new Set(notInMaster)];

      addCheck('Movement Quality', 'Movement names not in Masterlist',
        uniqueNotInMaster.length === 0 ? 'OK' : 'WARNING',
        uniqueNotInMaster.length === 0 ? 'All pending names matched'
          : uniqueNotInMaster.length + ' name(s) not found: ' + uniqueNotInMaster.slice(0, 5).join(', ') + (uniqueNotInMaster.length > 5 ? '…' : '')
      );
    }

    const noChangeFields = pendingRows.filter(r => {
      const typeOfMovement = normalize(getMovementValue(r, movementHeaders, 'Type of Movement'));
      return typeOfMovement !== 'ATTRITION' &&
             !getMovementValue(r, movementHeaders, 'New Department') &&
             !getMovementValue(r, movementHeaders, 'New Account') &&
             !getMovementValue(r, movementHeaders, 'New Supervisor') &&
             !getMovementValue(r, movementHeaders, 'New Job Title') &&
             !getMovementValue(r, movementHeaders, 'New Employment Status');
    }).length;
    addCheck('Movement Quality', 'Pending rows with no change fields',
      noChangeFields === 0 ? 'OK' : 'WARNING',
      noChangeFields === 0 ? 'None found' : noChangeFields + ' pending non-attrition row(s) have no New* fields filled — will produce no changes'
    );
  }

  const blocking = errors.length > 0;
  return finalize(ss, ts, checks, errors, warnings, stats, startTime, blocking);
}


// ── Write Health Report sheet + Pipeline Log ───────────────
function finalize(ss, ts, checks, errors, warnings, stats, startTime, blocking) {
  const durationSec   = Math.round((Date.now() - startTime) / 1000);
  const overallStatus = blocking ? 'error' : warnings.length > 0 ? 'warn' : 'ok';

  try {
    let report = ss.getSheetByName('Health Report');
    if (!report) report = ss.insertSheet('Health Report');
    report.clearContents();
    report.clearFormats();

    report.getRange(1, 1, 1, 4).setValues([['Masterlist Health Report', '', '', 'Last run: ' + ts]]);
    report.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    report.getRange(1, 4).setHorizontalAlignment('right').setFontColor('#6B7A99');

    const summaryRow = [
      'Active: ' + (stats.activeCount || '—'),
      'Inactive: ' + (stats.inactiveCount || '—'),
      'Pending movements: ' + (stats.pendingMovements !== undefined ? stats.pendingMovements : '—'),
      'History rows: ' + (stats.historyRows || '—')
    ];
    report.getRange(2, 1, 1, 4).setValues([summaryRow]);
    report.getRange(2, 1, 1, 4).setFontColor('#6B7A99').setFontStyle('italic');

    const bannerText   = blocking ? 'BLOCKED — fix errors before running'
                       : warnings.length > 0 ? 'Passed with warnings'
                       : 'All checks passed';
    const bannerColor  = blocking ? '#FDECEA' : warnings.length > 0 ? '#FFF7E6' : '#EAF7EC';
    const bannerFColor = blocking ? '#B02020' : warnings.length > 0 ? '#7A5000' : '#1B5E2A';
    report.getRange(3, 1, 1, 4).merge()
      .setValue(bannerText)
      .setBackground(bannerColor)
      .setFontColor(bannerFColor)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    report.getRange(4, 1, 1, 4).setValues([['Category', 'Check', 'Result', 'Detail']]);
    report.getRange(4, 1, 1, 4)
      .setBackground('#002B5C').setFontColor('#FFFFFF').setFontWeight('bold');

    const checkRows = checks.map(c => [c.category, c.checkName, c.result, c.detail]);
    if (checkRows.length > 0) {
      report.getRange(5, 1, checkRows.length, 4).setValues(checkRows);

      checks.forEach((c, i) => {
        const cell = report.getRange(5 + i, 3);
        if      (c.result === 'OK')      { cell.setBackground('#EAF7EC').setFontColor('#1B5E2A').setFontWeight('bold'); }
        else if (c.result === 'WARNING') { cell.setBackground('#FFF7E6').setFontColor('#7A5000').setFontWeight('bold'); }
        else if (c.result === 'ERROR')   { cell.setBackground('#FDECEA').setFontColor('#B02020').setFontWeight('bold'); }
        else                             { cell.setBackground('#F4F6FA').setFontColor('#6B7A99'); }
      });
    }

    report.setColumnWidth(1, 180);
    report.setColumnWidth(2, 260);
    report.setColumnWidth(3, 90);
    report.setColumnWidth(4, 420);

  } catch (e) {
    Logger.log('Health Report sheet write error (non-fatal): ' + e.message);
  }

  try {
    const logSS    = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    const logSheet = logSS.getSheetByName('Pipeline Log') || logSS.insertSheet('Pipeline Log');

    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 8).setValues([[
        'timestamp', 'stage', 'account',
        'movements_applied', 'history_rows_written',
        'duration_sec', 'status', 'notes'
      ]]);
      logSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    const noteText = blocking
      ? 'BLOCKED: ' + errors.slice(0, 2).join(' | ')
      : warnings.length > 0
        ? 'WARN: ' + warnings.slice(0, 2).join(' | ')
        : 'Health check passed — ' + checks.filter(c => c.result === 'OK').length + '/' + checks.length + ' checks OK';

    logSheet.appendRow([
      ts, 'healthcheck', LOG_ACCOUNT_NAME,
      stats.activeCount || 0,
      stats.historyRows || 0,
      durationSec,
      overallStatus,
      noteText
    ]);
  } catch (e) {
    Logger.log('Health check Pipeline Log error (non-fatal): ' + e.message);
  }

  return { blocking, errors, warnings, stats };
}


// ============================================================
//  PIPELINE LOGGER
// ============================================================

function logPipelineRun(movementsApplied, historyRowsWritten, durationSec, status, notes) {
  try {
    const ss    = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Pipeline Log') || ss.insertSheet('Pipeline Log');

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 8).setValues([[
        'timestamp', 'stage', 'account',
        'movements_applied', 'history_rows_written',
        'duration_sec', 'status', 'notes'
      ]]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      ts, 'appscript', LOG_ACCOUNT_NAME,
      movementsApplied, historyRowsWritten,
      durationSec, status, notes || ''
    ]);
  } catch (e) {
    Logger.log('Pipeline logger error (non-fatal): ' + e.message);
  }
}


// ============================================================
//  PROCESS MOVEMENTS
// ============================================================

function processMovements(master, movement, history, today, todayText, startTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const masterHeaders   = getHeaders(master);
  const movementHeaders = getHeaders(movement);

  const masterLastRow   = master.getLastRow();
  const movementLastRow = movement.getLastRow();

  const appliedChanges   = {};
  const warnings         = [];
  let   movementsApplied = 0;
  let   partial          = false;

  if (masterLastRow < 2 || movementLastRow < 2) {
    return { appliedChanges, movementsApplied, warnings, partial };
  }

  const masterData   = master.getRange(2, 1, masterLastRow - 1, master.getLastColumn()).getValues();
  const movementLastCol = movement.getLastColumn();
  const movementData = movement.getRange(2, 1, movementLastRow - 1, movementLastCol).getValues();

  const masterNameCol    = getCol(masterHeaders, 'Emp Name');
  const movementNameCol  = getCol(movementHeaders, 'Employee Name');
  const effectiveDateCol = getCol(movementHeaders, 'Effective Date');
  const timestampCol     = getCol(movementHeaders, 'Timestamp');

  const processedCol     = ensureColumn(movement, 'Processed');
  const processedDateCol = ensureColumn(movement, 'Processed Date');
  const processedNoteCol = ensureColumn(movement, 'Processed Note');
  const voidCol          = ensureColumn(movement, 'Void');

  const movementTrackingData = movement.getRange(2, 1, movementLastRow - 1, movement.getLastColumn()).getValues();
  const movementTrackingOut  = movementTrackingData.map(r => [...r]);

  const masterMap = {};
  masterData.forEach((row, index) => {
    const empName = normalize(row[masterNameCol]);
    if (empName) masterMap[empName] = { rowIndex: index + 2, row };
  });

  const eligibleMovements = [];

  for (let i = 0; i < movementData.length; i++) {
    const row      = movementData[i];
    const sheetRow = i + 2;

    const processed = normalize(movementTrackingOut[i][processedCol] || '');
    const isVoid    = normalize(movementTrackingOut[i][voidCol]      || '');

    if (processed === 'YES') continue;
    if (isVoid    === 'YES') continue;

    const effectiveDate = parseSheetDate(row[effectiveDateCol]);

    if (!effectiveDate) {
      movementTrackingOut[i][processedNoteCol] = 'Skipped: invalid or missing Effective Date';
      continue;
    }

    if (stripTime(effectiveDate) > stripTime(today)) {
      movementTrackingOut[i][processedNoteCol] = 'Pending Effective Date: ' +
        Utilities.formatDate(effectiveDate, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
      continue;
    }

    const empName = normalize(row[movementNameCol]);
    if (!empName) {
      movementTrackingOut[i][processedNoteCol] = 'Skipped: missing Employee Name';
      continue;
    }

    eligibleMovements.push({
      row, sheetRow, index: i, empName, effectiveDate,
      timestamp: timestampCol === -1 ? new Date(0) : (parseSheetDate(row[timestampCol]) || new Date(0))
    });
  }

  const latestMovementMap = {};
  eligibleMovements.forEach(item => {
    const key = item.empName + '|' + formatDateKey(item.effectiveDate);
    if (!latestMovementMap[key] || item.timestamp > latestMovementMap[key].timestamp) {
      latestMovementMap[key] = item;
    }
  });

  const latestMovements = Object.values(latestMovementMap)
    .sort((a, b) => a.effectiveDate - b.effectiveDate || a.timestamp - b.timestamp);

  for (let x = 0; x < latestMovements.length; x++) {
    if (remainingRuntimeMs(startTime) < 2 * 60 * 1000) {
      partial = true;
      break;
    }

    const item         = latestMovements[x];
    const empName      = item.empName;
    const masterRecord = masterMap[empName];

    if (!masterRecord) {
      movementTrackingOut[item.index][processedNoteCol] = 'Employee name not found in Masterlist';
      warnings.push('Not found in Masterlist: ' + empName);
      continue;
    }

    const changes = applyMovementToMasterlist(master, masterHeaders, item.row, movementHeaders, masterRecord.rowIndex);

    if (changes.length > 0) {
      item.changes = changes;
      backfillHistoryForMovement(history, movement, movementData, movementHeaders, master, masterHeaders, item, voidCol, startTime);
      if (isSameDate(item.effectiveDate, today)) {
        appliedChanges[empName] = changes.join(', ');
      }
      movementsApplied++;
    }

    movementTrackingOut[item.index][processedCol]     = 'Yes';
    movementTrackingOut[item.index][processedDateCol] = todayText;
    movementTrackingOut[item.index][processedNoteCol] = changes.length > 0
      ? changes.join(', ') + ' | History backfilled from Effective Date'
      : 'No applicable changes';
  }

  eligibleMovements.forEach(item => {
    const key    = item.empName + '|' + formatDateKey(item.effectiveDate);
    const latest = latestMovementMap[key];
    if (item.sheetRow !== latest.sheetRow) {
      movementTrackingOut[item.index][processedNoteCol] =
        'Skipped: older duplicate movement for same employee/effective date';
    }
  });

  movement.getRange(2, 1, movementLastRow - 1, movement.getLastColumn()).setValues(movementTrackingOut);

  return { appliedChanges, movementsApplied, warnings, partial };
}


// ============================================================
//  APPLY MOVEMENT TO MASTERLIST
// ============================================================

function applyMovementToMasterlist(master, masterHeaders, movementRow, movementHeaders, masterRowIndex) {
  const changes = [];

  const typeOfMovement      = getMovementValue(movementRow, movementHeaders, 'Type of Movement');
  const newDepartment       = getMovementValue(movementRow, movementHeaders, 'New Department');
  const newAccount          = getMovementValue(movementRow, movementHeaders, 'New Account');
  const newSupervisor       = getMovementValue(movementRow, movementHeaders, 'New Supervisor');
  const newJobTitle         = getMovementValue(movementRow, movementHeaders, 'New Job Title');
  const newEmploymentStatus = getMovementValue(movementRow, movementHeaders, 'New Employment Status') ||
                              (normalize(typeOfMovement) === 'ATTRITION' ? 'Inactive' : '');
  const effectiveDate       = getMovementValue(movementRow, movementHeaders, 'Effective Date');

  if (newDepartment) { setMasterValue(master, masterHeaders, masterRowIndex, 'Department', newDepartment); changes.push('Change of Department'); }
  if (newAccount)    { setMasterValue(master, masterHeaders, masterRowIndex, 'LOB / Account', newAccount); changes.push('Change of LOB / Account'); }

  if (newSupervisor) {
    setMasterValue(master, masterHeaders, masterRowIndex, 'Immediate Supervisor', newSupervisor);
    const managerName = getManagerOfNewSupervisor(master, masterHeaders, newSupervisor);
    if (managerName) setMasterValue(master, masterHeaders, masterRowIndex, 'Manager', managerName);
    changes.push('Change of Immediate Supervisor');
  }

  if (newJobTitle) { setMasterValue(master, masterHeaders, masterRowIndex, 'Job Title', newJobTitle); changes.push('Change of Job Title'); }

  if (newEmploymentStatus) {
    setMasterValue(master, masterHeaders, masterRowIndex, 'Employment Status', newEmploymentStatus);
    if (normalize(newEmploymentStatus) === 'INACTIVE' && effectiveDate) {
      setMasterValue(master, masterHeaders, masterRowIndex, 'Attrition Date', effectiveDate);
    }
    changes.push('Change of Employment Status');
  }

  return changes;
}


// ============================================================
//  BACKFILL HISTORY
// ============================================================

function backfillHistoryForMovement(history, movement, movementData, movementHeaders, master, masterHeaders, movementItem, voidCol, startTime) {
  const historyHeaders       = getHeaders(history);
  const historyNameCol       = getCol(historyHeaders, 'Emp Name');
  const historyDateCol       = getCol(historyHeaders, 'Date Generated');
  const historyChangeTypeCol = getCol(historyHeaders, 'Change Type');

  if (historyNameCol === -1 || historyDateCol === -1) return;

  const startDate = stripTime(movementItem.effectiveDate);
  const endDate   = getNextMovementEffectiveDate(movement, movementData, movementHeaders, movementItem, voidCol);

  const lastRow = history.getLastRow();
  if (lastRow < 2) return;

  const historyData    = history.getRange(2, 1, lastRow - 1, history.getLastColumn()).getValues();
  const updates        = buildHistoryBackfillUpdates(movementItem.row, movementHeaders, master, masterHeaders);
  const changeTypeText = movementItem.changes && movementItem.changes.length > 0 ? movementItem.changes.join(', ') : '';

  if (Object.keys(updates).length === 0 && !changeTypeText) return;

  const rowUpdates = [];
  historyData.forEach((historyRow, index) => {
    if (remainingRuntimeMs(startTime) < 90 * 1000) return;

    const rowNumber      = index + 2;
    const historyEmpName = normalize(historyRow[historyNameCol]);
    if (historyEmpName !== movementItem.empName) return;

    const historyDate = parseSheetDate(historyRow[historyDateCol]);
    if (!historyDate) return;

    const historyDateOnly = stripTime(historyDate);
    if (historyDateOnly < startDate) return;
    if (endDate && historyDateOnly >= endDate) return;

    const colUpdates = [];
    Object.keys(updates).forEach(headerName => {
      const col = getCol(historyHeaders, headerName);
      if (col !== -1) colUpdates.push({ col: col + 1, value: updates[headerName] });
    });
    if (historyChangeTypeCol !== -1 && changeTypeText && isSameDate(historyDateOnly, startDate)) {
      colUpdates.push({ col: historyChangeTypeCol + 1, value: changeTypeText });
    }
    if (colUpdates.length > 0) rowUpdates.push({ rowNumber, colUpdates });
  });

  rowUpdates.forEach(({ rowNumber, colUpdates }) => {
    colUpdates.forEach(({ col, value }) => history.getRange(rowNumber, col).setValue(value));
  });
}


// ============================================================
//  REFRESH TODAY'S HISTORY
// ============================================================

function refreshTodayHistory(master, history, todayText, weekStart, appliedChanges, startTime) {
  const masterHeaders       = getHeaders(master);
  const masterLastRow       = master.getLastRow();
  if (masterLastRow < 2) return 0;

  if (remainingRuntimeMs(startTime) < HISTORY_REFRESH_MIN_REMAINING_MS) return -1;

  deleteTodayHistoryRows(history, todayText, startTime);

  if (remainingRuntimeMs(startTime) < HISTORY_REFRESH_MIN_REMAINING_MS) return -1;

  const masterData          = master.getRange(2, 1, masterLastRow - 1, master.getLastColumn()).getValues();
  const employmentStatusCol = getCol(masterHeaders, 'Employment Status');
  const nameCol             = getCol(masterHeaders, 'Emp Name');

  const historyRows = masterData
    .filter(row => {
      const empName = normalize(row[nameCol]);
      const status  = normalize(row[employmentStatusCol]);
      return empName && status !== 'INACTIVE';
    })
    .map(row => {
      const empName    = normalize(row[nameCol]);
      const changeType = appliedChanges[empName] || getDefaultHistoryChangeType(history, empName);
      return [...row, todayText, changeType, weekStart];
    });

  if (historyRows.length === 0) return 0;

  history.getRange(history.getLastRow() + 1, 1, historyRows.length, historyRows[0].length).setValues(historyRows);
  return historyRows.length;
}


// ============================================================
//  SUPPORT FUNCTIONS
// ============================================================

function getNextMovementEffectiveDate(movement, movementData, movementHeaders, currentMovement, voidCol) {
  const movementNameCol  = getCol(movementHeaders, 'Employee Name');
  const effectiveDateCol = getCol(movementHeaders, 'Effective Date');
  let nextDate = null;

  movementData.forEach((row, index) => {
    const isVoid = normalize(row[voidCol] || '');
    if (isVoid === 'YES') return;

    const empName = normalize(row[movementNameCol]);
    if (empName !== currentMovement.empName) return;

    const effectiveDate = parseSheetDate(row[effectiveDateCol]);
    if (!effectiveDate) return;

    const effectiveDateOnly = stripTime(effectiveDate);
    const currentDateOnly   = stripTime(currentMovement.effectiveDate);

    if (effectiveDateOnly <= currentDateOnly) return;
    if (!nextDate || effectiveDateOnly < nextDate) nextDate = effectiveDateOnly;
  });

  return nextDate;
}

function buildHistoryBackfillUpdates(movementRow, movementHeaders, master, masterHeaders) {
  const updates = {};

  const typeOfMovement      = getMovementValue(movementRow, movementHeaders, 'Type of Movement');
  const newDepartment       = getMovementValue(movementRow, movementHeaders, 'New Department');
  const newAccount          = getMovementValue(movementRow, movementHeaders, 'New Account');
  const newSupervisor       = getMovementValue(movementRow, movementHeaders, 'New Supervisor');
  const newJobTitle         = getMovementValue(movementRow, movementHeaders, 'New Job Title');
  const newEmploymentStatus = getMovementValue(movementRow, movementHeaders, 'New Employment Status') ||
                              (normalize(typeOfMovement) === 'ATTRITION' ? 'Inactive' : '');
  const effectiveDate       = getMovementValue(movementRow, movementHeaders, 'Effective Date');

  if (newDepartment) updates['Department']          = newDepartment;
  if (newAccount)    updates['LOB / Account']       = newAccount;
  if (newJobTitle)   updates['Job Title']           = newJobTitle;

  if (newSupervisor) {
    updates['Immediate Supervisor'] = newSupervisor;
    const managerName = getManagerOfNewSupervisor(master, masterHeaders, newSupervisor);
    if (managerName) updates['Manager'] = managerName;
  }

  if (newEmploymentStatus) {
    updates['Employment Status'] = newEmploymentStatus;
    if (normalize(newEmploymentStatus) === 'INACTIVE' && effectiveDate) updates['Attrition Date'] = effectiveDate;
  }

  return updates;
}

function getManagerOfNewSupervisor(master, masterHeaders, newSupervisor) {
  const nameCol                = getCol(masterHeaders, 'Emp Name');
  const immediateSupervisorCol = getCol(masterHeaders, 'Immediate Supervisor');
  const managerCol             = getCol(masterHeaders, 'Manager');

  if (nameCol === -1) return '';
  const lastRow = master.getLastRow();
  if (lastRow < 2) return '';

  const data          = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
  const supervisorRow = data.find(row => normalizeName(row[nameCol]) === normalizeName(newSupervisor));

  if (!supervisorRow) return '';
  if (immediateSupervisorCol !== -1 && supervisorRow[immediateSupervisorCol]) return supervisorRow[immediateSupervisorCol];
  if (managerCol !== -1 && supervisorRow[managerCol]) return supervisorRow[managerCol];
  return '';
}

function setupHistorySheet(master, history) {
  if (history.getLastRow() > 0) return;
  const masterHeaders  = getHeaders(master);
  const historyHeaders = [...masterHeaders, 'Date Generated', 'Change Type', 'Week'];
  history.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
}

function setupMovementTrackingColumns(movement) {
  ensureColumn(movement, 'Processed');
  ensureColumn(movement, 'Processed Date');
  ensureColumn(movement, 'Processed Note');
  ensureColumn(movement, 'Void');
}

function deleteTodayHistoryRows(history, todayText, startTime) {
  const headers = getHeaders(history);
  const dateCol = getCol(headers, 'Date Generated');
  const lastRow = history.getLastRow();
  if (lastRow < 2 || dateCol === -1) return;

  const dateValues = history.getRange(2, dateCol + 1, lastRow - 1, 1).getValues();
  for (let row = lastRow; row >= 2; row--) {
    if (remainingRuntimeMs(startTime) < 90 * 1000) return;
    const value     = dateValues[row - 2][0];
    const valueText = value instanceof Date
      ? Utilities.formatDate(value, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy')
      : String(value).trim();
    if (valueText === todayText) history.deleteRow(row);
  }
}

function getDefaultHistoryChangeType(history, empName) {
  const headers = getHeaders(history);
  const nameCol = getCol(headers, 'Emp Name');
  const lastRow = history.getLastRow();
  if (lastRow < 2 || nameCol === -1) return 'New Record';

  const data   = history.getRange(2, 1, lastRow - 1, history.getLastColumn()).getValues();
  const exists = data.some(row => normalize(row[nameCol]) === empName);
  return exists ? 'No Change' : 'New Record';
}

function getMovementValue(row, headers, headerName) {
  const col = getCol(headers, headerName);
  if (col === -1) return '';
  const value   = row[col];
  const cleaned = String(value).trim();
  if (cleaned === '' || cleaned.toUpperCase() === 'N/A' || cleaned.toUpperCase() === 'NA' || cleaned === '-') return '';
  return value;
}

function setMasterValue(sheet, headers, rowIndex, headerName, value) {
  const col = getCol(headers, headerName);
  if (col === -1) return;
  sheet.getRange(rowIndex, col + 1).setValue(value);
}

function ensureColumn(sheet, headerName) {
  const headers     = getHeaders(sheet);
  const existingCol = getCol(headers, headerName);
  if (existingCol !== -1) return existingCol;
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(headerName);
  return newCol - 1;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getCol(headers, headerName) {
  return headers.findIndex(header => normalize(header) === normalize(headerName));
}

function normalize(value) {
  return String(value).trim().replace(/\s+/g, ' ').toUpperCase();
}

function normalizeName(value) {
  return String(value).trim().replace(/\./g, '').replace(/\s+/g, ' ').toUpperCase();
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function isSameDate(date1, date2) {
  return stripTime(date1).getTime() === stripTime(date2).getTime();
}

function formatDateKey(date) {
  return Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function parseSheetDate(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function getWeekStart(dateText) {
  const date = new Date(dateText);
  const day  = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy');
}

function remainingRuntimeMs(startTime) {
  return MAX_RUNTIME_MS - (Date.now() - startTime);
}

function appendNote(existing, note) {
  return existing ? existing + ' | ' + note : note;
}
