/**
 * ============================================================
 *  SHIFT SYNC — wide (pivoted) tblShift
 * ============================================================
 *  Layout:
 *    Row 1: Emp ID | yyyy-MM-dd | yyyy-MM-dd | ...
 *    Row 2+: empId | shift      | shift      | ...
 *
 *  Every run (manual or trigger) does ALL of the following in memory:
 *
 *  A. RETENTION (columns)
 *     Keep from the 1st of the current month, 12 months ago.
 *     e.g. today Aug 2026 → keep from 2025-08-01 (drop through 2025-07-31).
 *     Date columns older than the cutoff are never written back.
 *
 *  B. CSV RANGE REPLACE (columns)
 *     Dates in [csvMin, csvMax] are removed from the existing grid, then
 *     the CSV is unpivoted and written as those date columns again.
 *     Employees are matched by Emp ID.
 *
 *  C. EMPTY EMPLOYEE ROWS
 *     After A + B, any employee with zero remaining shift cells is dropped.
 *
 *  D. WRITE
 *     Sheet is recreated (avoids ghost 300k rows / cell-limit errors),
 *     then the compact wide matrix is written in one pass.
 * ============================================================
 */

var EMP_SHIFT_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var EMP_SHIFT_CSV_NAME = 'Employee Shift.csv';
var SHIFT_WRITE_CHUNK_ROWS = 500;

function processShiftFiles() {
  return syncShiftsFromCsv();
}

function syncShiftsFromCsv() {
  var started = new Date().getTime();
  var cutoffStr = ymd_(shiftRetentionCutoff_());

  // ---- 1. CSV (long → per-emp map; min/max for range replace) ----
  var csv = loadEmployeeShiftCsvWide_();
  if (!csv.success) {
    return { success: false, message: csv.message };
  }
  Logger.log('CSV: ' + csv.rowCount + ' cells, emps ' + csv.empCount +
    ', dates ' + csv.minDate + ' → ' + csv.maxDate +
    ' (' + (new Date().getTime() - started) + ' ms)');

  // ---- 2. Existing sheet → grid (already drops < cutoff and CSV range) ----
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  if (!sheet) sheet = ss.insertSheet('tblShift');

  var grid = loadExistingShiftGrid_(sheet, cutoffStr, csv.minDate, csv.maxDate);
  Logger.log('Existing grid emps (after retention + CSV-range drop): ' +
    Object.keys(grid).length + ' (' + (new Date().getTime() - started) + ' ms)');

  // ---- 3. Unpivot CSV into grid (Emp ID mapped; CSV wins for its dates) ----
  var empIds = Object.keys(csv.byEmp);
  for (var i = 0; i < empIds.length; i++) {
    var empId = empIds[i];
    if (!grid[empId]) grid[empId] = {};
    var dates = csv.byEmp[empId];
    var dk = Object.keys(dates);
    for (var j = 0; j < dk.length; j++) {
      // CSV rows older than cutoff were already skipped at load
      grid[empId][dk[j]] = dates[dk[j]];
    }
  }

  // ---- 4. Build date column list (retention only) + prune empty emps ----
  var allDates = {};
  var activeEmps = [];
  var prunedEmps = 0;
  var gEmps = Object.keys(grid);

  for (var e = 0; e < gEmps.length; e++) {
    var id = gEmps[e];
    var m = grid[id];
    var kept = {};
    var hasData = false;
    var ds = Object.keys(m);
    for (var d = 0; d < ds.length; d++) {
      var dStr = ds[d];
      if (dStr < cutoffStr) continue; // retention: drop old columns
      var code = m[dStr];
      if (!code) continue;
      kept[dStr] = code;
      allDates[dStr] = true;
      hasData = true;
    }
    if (hasData) {
      grid[id] = kept;
      activeEmps.push(id);
    } else {
      delete grid[id]; // no shifts left in retention window → drop row
      prunedEmps++;
    }
  }

  var dateList = Object.keys(allDates).sort();
  activeEmps.sort();

  var headers = ['Emp ID'].concat(dateList);
  var out = [headers];
  for (var r = 0; r < activeEmps.length; r++) {
    var eid = activeEmps[r];
    var row = [eid];
    var map = grid[eid];
    for (var c = 0; c < dateList.length; c++) {
      row.push(map[dateList[c]] || '');
    }
    out.push(row);
  }

  Logger.log('Wide matrix: ' + activeEmps.length + ' emps × ' + dateList.length +
    ' dates (pruned empty emps: ' + prunedEmps + ') (' +
    (new Date().getTime() - started) + ' ms)');

  // ---- 5. Recreate sheet + write ----
  sheet = resetTblShiftSheet_(ss, sheet);
  writeWideChunked_(sheet, out);
  SpreadsheetApp.flush();

  try {
    if (typeof cacheClearAll_ === 'function') cacheClearAll_();
    CacheService.getScriptCache().remove('emp_map');
  } catch (e2) {}

  var ms = new Date().getTime() - started;
  var msg = 'Shift wide-sync OK in ' + ms + ' ms. ' +
    activeEmps.length + ' employees, ' + dateList.length + ' date columns' +
    (dateList.length ? ' (' + dateList[0] + ' → ' + dateList[dateList.length - 1] + ')' : '') +
    '. CSV range ' + csv.minDate + '→' + csv.maxDate + ' replaced. ' +
    'Cutoff ' + cutoffStr + '. Empty emp rows pruned: ' + prunedEmps + '.';
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    employees: activeEmps.length,
    dates: dateList.length,
    dateFrom: dateList.length ? dateList[0] : null,
    dateTo: dateList.length ? dateList[dateList.length - 1] : null,
    csvMin: csv.minDate,
    csvMax: csv.maxDate,
    cutoff: cutoffStr,
    prunedEmps: prunedEmps,
    elapsedMs: ms
  };
}

/**
 * Replace tblShift with a fresh sheet so leftover long-format rows cannot
 * expand into millions of cells when wide columns are added.
 */
function resetTblShiftSheet_(ss, sheet) {
  var name = 'tblShift';
  var idx = sheet.getIndex();

  try {
    ss.deleteSheet(sheet);
    var fresh = ss.insertSheet(name, Math.max(0, idx - 1));
    Logger.log('tblShift recreated at index ' + idx);
    return fresh;
  } catch (err) {
    Logger.log('deleteSheet failed (' + err.message + '); trimming rows instead');
    sheet.clearContents();
    sheet.clearFormats();
    var maxRows = sheet.getMaxRows();
    if (maxRows > 2000) sheet.deleteRows(2001, maxRows - 2000);
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 400) sheet.deleteColumns(401, maxCols - 400);
    return sheet;
  }
}

/** 1st of current month, 12 months ago. */
function shiftRetentionCutoff_() {
  var now = new Date();
  return new Date(now.getFullYear() - 1, now.getMonth(), 1);
}

function ymd_(d) {
  return d.getFullYear() + '-' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
    ('0' + d.getDate()).slice(-2);
}

function ymdFromAny_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) return ymd_(val);
  var s = String(val).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return ymd_(d);
}

/**
 * Existing tblShift → grid.
 * Always applies:
 *   - drop dates < cutoffStr (retention)
 *   - drop dates in [csvMin, csvMax] (will be refilled from CSV)
 * Supports wide and legacy long layouts.
 */
function loadExistingShiftGrid_(sheet, cutoffStr, csvMin, csvMax) {
  var grid = {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return grid;

  var h0 = String(data[0][0] || '').trim().toLowerCase();
  var h1 = String(data[0][1] || '').trim();

  var isWide = /^\d{4}-\d{2}-\d{2}/.test(h1) ||
    (h0.indexOf('emp') === 0 && h1.toLowerCase() !== 'date');

  if (isWide) {
    var dateHeaders = [];
    for (var c = 1; c < data[0].length; c++) {
      dateHeaders.push(ymdFromAny_(data[0][c]));
    }
    for (var r = 1; r < data.length; r++) {
      var empId = String(data[r][0] || '').trim().toUpperCase();
      if (!empId) continue;
      if (!grid[empId]) grid[empId] = {};
      for (var c2 = 1; c2 < data[r].length; c2++) {
        var dStr = dateHeaders[c2 - 1];
        if (!dStr) continue;
        if (dStr < cutoffStr) continue;
        if (csvMin && csvMax && dStr >= csvMin && dStr <= csvMax) continue;
        var sh = String(data[r][c2] || '').trim().toUpperCase();
        if (sh) grid[empId][dStr] = sh;
      }
    }
    return grid;
  }

  // Legacy long: Emp ID | Date | Shift
  for (var i = 1; i < data.length; i++) {
    var emp = String(data[i][0] || '').trim().toUpperCase();
    if (!emp) continue;
    var dStr2 = ymdFromAny_(data[i][1]);
    if (!dStr2) continue;
    if (dStr2 < cutoffStr) continue;
    if (csvMin && csvMax && dStr2 >= csvMin && dStr2 <= csvMax) continue;
    var code = String(data[i][2] || '').trim().toUpperCase();
    if (!code) continue;
    if (!grid[emp]) grid[emp] = {};
    grid[emp][dStr2] = code;
  }
  return grid;
}

/**
 * CSV long rows → { byEmp, minDate, maxDate }.
 * Skips shifts older than retention cutoff.
 */
function loadEmployeeShiftCsvWide_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SHIFT_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SHIFT_CSV_NAME);
    if (!files.hasNext()) {
      return { success: false, message: 'CSV not found: ' + EMP_SHIFT_CSV_NAME };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var parsed = Utilities.parseCsv(file.getBlob().getDataAsString());
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'Shift CSV empty' };
    }

    var rawH = parsed[0].map(function (h) { return String(h || '').trim(); });
    var headers = rawH.map(function (h) {
      var m = h.match(/^Shift\[(.+)\]$/i);
      return m ? m[1].trim() : h;
    });

    var find = function (names) {
      for (var n = 0; n < names.length; n++) {
        var w = names[n].toLowerCase();
        for (var k = 0; k < headers.length; k++) {
          if (String(headers[k]).toLowerCase() === w) return k;
        }
      }
      return -1;
    };

    var iEmp = find(['Emp No', 'Emp ID', 'Employee ID']);
    var iDate = find(['Tr Date', 'Date', 'Shift Date']);
    var iShift = find(['Final Shift', 'Shift', 'Shift Code', 'Code']);
    if (iEmp < 0 || iDate < 0 || iShift < 0) {
      return { success: false, message: 'CSV columns missing. Found: ' + headers.join(', ') };
    }

    var cutoffStr = ymd_(shiftRetentionCutoff_());
    var byEmp = {};
    var minDate = null;
    var maxDate = null;
    var rowCount = 0;

    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row || !row.length) continue;
      var empId = String(row[iEmp] || '').trim().toUpperCase();
      if (!empId) continue;

      var rawD = String(row[iDate] || '').trim();
      var dStr = null;
      if (rawD.length >= 10 && rawD.charAt(4) === '-' && rawD.charAt(7) === '-') {
        dStr = rawD.substring(0, 10);
      } else {
        dStr = ymdFromAny_(rawD);
      }
      if (!dStr || dStr < cutoffStr) continue;

      var shift = String(row[iShift] || '').trim().toUpperCase();
      if (!shift) continue;

      if (!byEmp[empId]) byEmp[empId] = {};
      byEmp[empId][dStr] = shift;
      rowCount++;
      if (!minDate || dStr < minDate) minDate = dStr;
      if (!maxDate || dStr > maxDate) maxDate = dStr;
    }

    return {
      success: true,
      byEmp: byEmp,
      minDate: minDate,
      maxDate: maxDate,
      rowCount: rowCount,
      empCount: Object.keys(byEmp).length
    };
  } catch (err) {
    return { success: false, message: 'CSV error: ' + err.message };
  }
}

function writeWideChunked_(sheet, rows) {
  if (!rows || !rows.length) return;
  var cols = rows[0].length;
  var need = cols - sheet.getMaxColumns();
  if (need > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), need);

  var chunk = SHIFT_WRITE_CHUNK_ROWS;
  for (var i = 0; i < rows.length; i += chunk) {
    var part = rows.slice(i, i + chunk);
    sheet.getRange(i + 1, 1, part.length, cols).setValues(part);
    if (i + chunk < rows.length) SpreadsheetApp.flush();
  }
}

function setupDailyShiftTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if ((fn === 'processShiftFiles' || fn === 'syncShiftsFromCsv') &&
        t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processShiftFiles')
    .timeBased()
    .atHour(1)
    .nearMinute(0)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger: processShiftFiles ~1:00 AM');
}
