/**
 * ============================================================
 *  SHIFT SYNC — wide (pivoted) tblShift
 * ============================================================
 *  Layout:
 *    Row 1: Emp ID | 2026-07-01 | 2026-07-02 | ...   (ISO date TEXT only)
 *    Row 2+: empId | G          | O          | ...
 *
 *  CRITICAL: Date keys stay as ISO text (yyyy-MM-dd) end-to-end.
 *  Never construct JavaScript Date for shift keys — no locale
 *  day/month swap, no timezone shift. CSV "2026-07-01T00:00:00"
 *  → key "2026-07-01" by string slice only. Pivot and write back.
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
  var cutoffStr = isoTodayCutoff_();
  Logger.log('=== processShiftFiles START === cutoff(ISO)=' + cutoffStr);

  var csv = loadEmployeeShiftCsvWide_();
  if (!csv.success) {
    Logger.log('CSV load failed: ' + csv.message);
    return { success: false, message: csv.message };
  }
  Logger.log('CSV: ' + csv.rowCount + ' cells, emps ' + csv.empCount +
    ', ISO dates ' + csv.minDate + ' → ' + csv.maxDate +
    ' (sample keys: ' + (csv.sampleKeys || []).join(', ') + ')' +
    ' (' + (new Date().getTime() - started) + ' ms)');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  if (!sheet) sheet = ss.insertSheet('tblShift');

  var grid = loadExistingShiftGrid_(sheet, cutoffStr, csv.minDate, csv.maxDate);
  Logger.log('Existing grid emps (after retention + CSV-range drop): ' +
    Object.keys(grid).length + ' (' + (new Date().getTime() - started) + ' ms)');

  var empIds = Object.keys(csv.byEmp);
  for (var i = 0; i < empIds.length; i++) {
    var empId = empIds[i];
    if (!grid[empId]) grid[empId] = {};
    var dates = csv.byEmp[empId];
    var dk = Object.keys(dates);
    for (var j = 0; j < dk.length; j++) {
      grid[empId][dk[j]] = dates[dk[j]];
    }
  }

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
      if (!isIsoDateKey_(dStr)) continue;
      if (dStr < cutoffStr) continue;
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
      delete grid[id];
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
    ' ISO date cols. First/last: ' +
    (dateList.length ? dateList[0] + ' → ' + dateList[dateList.length - 1] : '(none)') +
    '. Sample: ' + dateList.slice(0, 6).join(', '));

  sheet = resetTblShiftSheet_(ss, sheet);
  writeWideAsIsoText_(sheet, out);
  SpreadsheetApp.flush();

  try {
    if (typeof cacheClearAll_ === 'function') cacheClearAll_();
    CacheService.getScriptCache().remove('emp_map');
  } catch (e2) {}

  var ms = new Date().getTime() - started;
  var msg = 'Shift wide-sync OK in ' + ms + ' ms. ' +
    activeEmps.length + ' employees, ' + dateList.length + ' ISO date columns' +
    (dateList.length ? ' (' + dateList[0] + ' → ' + dateList[dateList.length - 1] + ')' : '') +
    '. CSV range ' + csv.minDate + '→' + csv.maxDate + ' replaced. ' +
    'Cutoff ' + cutoffStr + '. Empty emp rows pruned: ' + prunedEmps + '.';
  Logger.log('=== processShiftFiles END === ' + msg);
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

function isIsoDateKey_(s) {
  return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
}

/**
 * Extract yyyy-MM-dd by STRING only.
 * "2026-07-01T00:00:00" → "2026-07-01". Never new Date() for CSV keys.
 */
function isoDateKeyFromText_(val) {
  if (val === null || val === undefined || val === '') return null;

  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
      var y = val.getFullYear();
      var m = val.getMonth() + 1;
      var day = val.getDate();
      return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + day).slice(-2);
    }
  }

  var s = String(val).trim();
  if (s.charAt(0) === "'") s = s.substring(1);

  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  return null;
}

function isoTodayCutoff_() {
  var now = new Date();
  var y = now.getFullYear() - 1;
  var m = now.getMonth() + 1;
  return y + '-' + ('0' + m).slice(-2) + '-01';
}

function resetTblShiftSheet_(ss, sheet) {
  var name = 'tblShift';
  var idx = sheet.getIndex();
  try {
    ss.deleteSheet(sheet);
    var fresh = ss.insertSheet(name, Math.max(0, idx - 1));
    Logger.log('tblShift recreated at index ' + idx);
    return fresh;
  } catch (err) {
    Logger.log('deleteSheet failed (' + err.message + '); trimming instead');
    sheet.clearContents();
    sheet.clearFormats();
    var maxRows = sheet.getMaxRows();
    if (maxRows > 2000) sheet.deleteRows(2001, maxRows - 2000);
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 400) sheet.deleteColumns(401, maxCols - 400);
    return sheet;
  }
}

function writeWideAsIsoText_(sheet, rows) {
  if (!rows || !rows.length) return;
  var cols = rows[0].length;
  var need = cols - sheet.getMaxColumns();
  if (need > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), need);

  if (cols > 1) {
    sheet.getRange(1, 2, 1, cols - 1).setNumberFormat('@');
  }

  for (var c = 1; c < rows[0].length; c++) {
    var key = isoDateKeyFromText_(rows[0][c]) || String(rows[0][c] || '');
    rows[0][c] = key;
  }

  var chunk = SHIFT_WRITE_CHUNK_ROWS;
  for (var i = 0; i < rows.length; i += chunk) {
    var part = rows.slice(i, i + chunk);
    sheet.getRange(i + 1, 1, part.length, cols).setValues(part);
    if (i === 0 && cols > 1) {
      sheet.getRange(1, 2, 1, cols - 1).setNumberFormat('@');
    }
    if (i + chunk < rows.length) SpreadsheetApp.flush();
  }

  try {
    var shown = sheet.getRange(1, 2, 1, Math.min(6, cols - 1)).getDisplayValues()[0];
    Logger.log('Header display sample after write: ' + shown.join(' | '));
    var raw = sheet.getRange(1, 2, 1, Math.min(6, cols - 1)).getValues()[0];
    Logger.log('Header raw typeof sample: ' + raw.map(function (v) {
      return (v instanceof Date ? 'Date(' + isoDateKeyFromText_(v) + ')' : typeof v + ':' + v);
    }).join(', '));
  } catch (e) {}
}

function loadExistingShiftGrid_(sheet, cutoffStr, csvMin, csvMax) {
  var grid = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return grid;

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var display = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  var h0 = String(data[0][0] || '').trim().toLowerCase();
  var h1Key = isoDateKeyFromText_(display[1]) || isoDateKeyFromText_(data[0][1]);
  var isWide = isIsoDateKey_(h1Key) ||
    (h0.indexOf('emp') === 0 && String(display[1] || data[0][1] || '').toLowerCase() !== 'date');

  if (isWide) {
    var dateHeaders = [];
    for (var c = 1; c < lastCol; c++) {
      var key = isoDateKeyFromText_(display[c]) || isoDateKeyFromText_(data[0][c]);
      dateHeaders.push(key);
    }
    Logger.log('loadExistingShiftGrid_ wide headers sample: ' +
      dateHeaders.filter(Boolean).slice(0, 5).join(', '));

    for (var r = 1; r < data.length; r++) {
      var empId = String(data[r][0] || '').trim().toUpperCase();
      if (!empId) continue;
      if (!grid[empId]) grid[empId] = {};
      for (var c2 = 1; c2 < data[r].length; c2++) {
        var dStr = dateHeaders[c2 - 1];
        if (!dStr || !isIsoDateKey_(dStr)) continue;
        if (dStr < cutoffStr) continue;
        if (csvMin && csvMax && dStr >= csvMin && dStr <= csvMax) continue;
        var sh = String(data[r][c2] || '').trim().toUpperCase();
        if (sh) grid[empId][dStr] = sh;
      }
    }
    return grid;
  }

  for (var i = 1; i < data.length; i++) {
    var emp = String(data[i][0] || '').trim().toUpperCase();
    if (!emp) continue;
    var dStr2 = isoDateKeyFromText_(data[i][1]);
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

function loadEmployeeShiftCsvWide_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SHIFT_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SHIFT_CSV_NAME);
    var file = null;
    if (files.hasNext()) {
      file = files.next();
      while (files.hasNext()) {
        var f2 = files.next();
        if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
      }
    } else {
      var all = folder.getFiles();
      while (all.hasNext()) {
        var f = all.next();
        if (/employee\s*shift.*\.csv$/i.test(f.getName())) {
          if (!file || f.getLastUpdated() > file.getLastUpdated()) file = f;
        }
      }
    }
    if (!file) {
      return { success: false, message: 'CSV not found: ' + EMP_SHIFT_CSV_NAME };
    }

    var text = file.getBlob().getDataAsString().replace(/^\uFEFF/, '');
    var parsed = Utilities.parseCsv(text);
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

    var cutoffStr = isoTodayCutoff_();
    var byEmp = {};
    var minDate = null;
    var maxDate = null;
    var rowCount = 0;
    var badDates = 0;
    var sampleKeys = [];

    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row || !row.length) continue;
      var empId = String(row[iEmp] || '').trim().toUpperCase();
      if (!empId) continue;

      var rawD = String(row[iDate] || '').trim();
      var dStr = isoDateKeyFromText_(rawD);
      if (!dStr) {
        badDates++;
        if (badDates <= 5) Logger.log('Bad shift date (not ISO): "' + rawD + '" emp ' + empId);
        continue;
      }
      if (dStr < cutoffStr) continue;

      var shift = String(row[iShift] || '').trim().toUpperCase();
      if (!shift) continue;

      if (!byEmp[empId]) byEmp[empId] = {};
      byEmp[empId][dStr] = shift;
      rowCount++;
      if (!minDate || dStr < minDate) minDate = dStr;
      if (!maxDate || dStr > maxDate) maxDate = dStr;
      if (sampleKeys.length < 8 && sampleKeys.indexOf(dStr) < 0) sampleKeys.push(dStr);
    }

    sampleKeys.sort();
    Logger.log('CSV parse (string ISO keys): ' + rowCount + ' cells, badDates=' + badDates +
      ', range ' + minDate + ' → ' + maxDate + ', samples ' + sampleKeys.join(', '));

    return {
      success: true,
      byEmp: byEmp,
      minDate: minDate,
      maxDate: maxDate,
      rowCount: rowCount,
      empCount: Object.keys(byEmp).length,
      badDates: badDates,
      sampleKeys: sampleKeys
    };
  } catch (err) {
    return { success: false, message: 'CSV error: ' + err.message };
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
