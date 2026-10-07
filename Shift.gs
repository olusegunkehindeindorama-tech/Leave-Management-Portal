/**
 * ============================================================
 *  SHIFT SYNC — wide (pivoted) tblShift
 * ============================================================
 *  SOURCE OF TRUTH = Employee Shift.csv only.
 *
 *  - Every date key is PLAIN TEXT "yyyy-MM-dd" (ISO).
 *  - Never new Date() for shift keys. Never locale parse.
 *  - tblShift is fully rebuilt from the CSV each run.
 *    Old sheet columns (e.g. wrong Nov/Dec from past bugs) are discarded.
 *  - CSV "2026-07-01T00:00:00" → header "2026-07-01" by string slice only.
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
  Logger.log('=== processShiftFiles START (CSV-only, text ISO keys) ===');

  var csv = loadEmployeeShiftCsvWide_();
  if (!csv.success) {
    Logger.log('CSV load failed: ' + csv.message);
    return { success: false, message: csv.message };
  }

  Logger.log(
    'CSV loaded: ' + csv.rowCount + ' shift cells, ' + csv.empCount +
    ' employees, ISO dates ' + csv.minDate + ' → ' + csv.maxDate +
    ' (' + csv.dateCount + ' distinct date columns). Samples: ' +
    (csv.sampleKeys || []).join(', ')
  );

  var dateList = csv.dateList.slice();
  var empList = Object.keys(csv.byEmp).sort();

  var headers = ['Emp ID'].concat(dateList);
  var out = [headers];

  for (var r = 0; r < empList.length; r++) {
    var empId = empList[r];
    var map = csv.byEmp[empId];
    var row = [empId];
    for (var c = 0; c < dateList.length; c++) {
      row.push(map[dateList[c]] || '');
    }
    out.push(row);
  }

  Logger.log(
    'Wide matrix from CSV only: ' + empList.length + ' emps × ' +
    dateList.length + ' text ISO columns. First/last: ' +
    (dateList.length ? dateList[0] + ' → ' + dateList[dateList.length - 1] : '(none)')
  );

  for (var i = 0; i < dateList.length; i++) {
    if (!isIsoDateKey_(dateList[i])) {
      var msg = 'Abort: non-ISO date key in memory: "' + dateList[i] + '"';
      Logger.log(msg);
      return { success: false, message: msg };
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  if (!sheet) sheet = ss.insertSheet('tblShift');

  sheet = resetTblShiftSheet_(ss, sheet);
  writeWideAsText_(sheet, out);
  SpreadsheetApp.flush();

  try {
    if (typeof cacheClearAll_ === 'function') cacheClearAll_();
    CacheService.getScriptCache().remove('emp_map');
  } catch (e2) {}

  var ms = new Date().getTime() - started;
  var msg =
    'Shift rebuild OK in ' + ms + ' ms (CSV-only, text ISO). ' +
    empList.length + ' employees, ' + dateList.length + ' date columns (' +
    (dateList.length ? dateList[0] + ' → ' + dateList[dateList.length - 1] : 'none') +
    '). Previous tblShift date columns discarded.';
  Logger.log('=== processShiftFiles END === ' + msg);

  return {
    success: true,
    message: msg,
    employees: empList.length,
    dates: dateList.length,
    dateFrom: dateList.length ? dateList[0] : null,
    dateTo: dateList.length ? dateList[dateList.length - 1] : null,
    csvMin: csv.minDate,
    csvMax: csv.maxDate,
    elapsedMs: ms
  };
}

function isIsoDateKey_(s) {
  return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function isoDateKeyFromText_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  if (s.charAt(0) === "'") s = s.substring(1);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  var month = Number(m[2]);
  var day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return m[1] + '-' + m[2] + '-' + m[3];
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
    Logger.log('Using CSV file: ' + file.getName() + ' (updated ' + file.getLastUpdated() + ')');

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
      return {
        success: false,
        message: 'CSV columns missing. Found: ' + headers.join(', ')
      };
    }

    var byEmp = {};
    var allDates = {};
    var minDate = null;
    var maxDate = null;
    var rowCount = 0;
    var badDates = 0;
    var sampleRaw = [];

    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row || !row.length) continue;

      var empId = String(row[iEmp] || '').trim().toUpperCase();
      if (!empId) continue;

      var rawD = String(row[iDate] || '').trim();
      var dStr = isoDateKeyFromText_(rawD);
      if (!dStr) {
        badDates++;
        if (sampleRaw.length < 5) {
          sampleRaw.push(rawD);
          Logger.log('Skip non-ISO date: "' + rawD + '" emp=' + empId);
        }
        continue;
      }

      var shift = String(row[iShift] || '').trim().toUpperCase();
      if (!shift) continue;

      if (!byEmp[empId]) byEmp[empId] = {};
      byEmp[empId][dStr] = shift;
      allDates[dStr] = true;
      rowCount++;

      if (!minDate || dStr < minDate) minDate = dStr;
      if (!maxDate || dStr > maxDate) maxDate = dStr;
    }

    var dateList = Object.keys(allDates).sort();
    var sampleKeys = dateList.slice(0, 8);

    Logger.log(
      'CSV text-pivot: cells=' + rowCount + ' badDates=' + badDates +
      ' distinctDays=' + dateList.length +
      ' range=' + minDate + '→' + maxDate
    );

    var novDec = dateList.filter(function (d) {
      return /-11-|-12-/.test(d);
    });
    if (novDec.length) {
      Logger.log('NOTE: CSV itself contains Nov/Dec ISO keys: ' +
        novDec.slice(0, 10).join(', ') +
        (novDec.length > 10 ? ' … +' + (novDec.length - 10) + ' more' : ''));
    } else {
      Logger.log('CSV has no November/December keys (good).');
    }

    return {
      success: true,
      byEmp: byEmp,
      dateList: dateList,
      dateCount: dateList.length,
      minDate: minDate,
      maxDate: maxDate,
      rowCount: rowCount,
      empCount: Object.keys(byEmp).length,
      badDates: badDates,
      sampleKeys: sampleKeys
    };
  } catch (err) {
    Logger.log('CSV error: ' + err.message + '\n' + err.stack);
    return { success: false, message: 'CSV error: ' + err.message };
  }
}

function resetTblShiftSheet_(ss, sheet) {
  var name = 'tblShift';
  var idx = sheet.getIndex();
  try {
    ss.deleteSheet(sheet);
    var fresh = ss.insertSheet(name, Math.max(0, idx - 1));
    Logger.log('tblShift recreated (old columns gone)');
    return fresh;
  } catch (err) {
    Logger.log('deleteSheet failed (' + err.message + '); clear instead');
    sheet.clearContents();
    sheet.clearFormats();
    var maxRows = sheet.getMaxRows();
    if (maxRows > 2000) sheet.deleteRows(2001, maxRows - 2000);
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 400) sheet.deleteColumns(401, maxCols - 400);
    return sheet;
  }
}

function writeWideAsText_(sheet, rows) {
  if (!rows || !rows.length) return;
  var cols = rows[0].length;
  var need = cols - sheet.getMaxColumns();
  if (need > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), need);

  if (cols > 1) {
    sheet.getRange(1, 2, 1, cols - 1).setNumberFormat('@');
  }

  for (var c = 1; c < rows[0].length; c++) {
    rows[0][c] = String(rows[0][c] || '');
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
    var n = Math.min(8, cols - 1);
    if (n > 0) {
      var shown = sheet.getRange(1, 2, 1, n).getDisplayValues()[0];
      var raw = sheet.getRange(1, 2, 1, n).getValues()[0];
      Logger.log('Header DISPLAY: ' + shown.join(' | '));
      Logger.log('Header RAW types: ' + raw.map(function (v) {
        if (v instanceof Date) return 'Date!';
        return typeof v + '=' + v;
      }).join(', '));
    }
  } catch (e) {
    Logger.log('Header verify error: ' + e.message);
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
