/**
 * ============================================================
 *  SHIFT SYNC — Google Apps Script (trigger-friendly)
 * ============================================================
 *  Source: "Employee Shift.csv" in folder EMP_SHIFT_FOLDER_ID
 *  Target: sheet tblShift  [Emp ID | Date | Shift]
 *
 *  Fast path (avoids 6-min timeout on ~100k CSV rows):
 *   1. Load CSV → rows + set of dates present in CSV
 *   2. Load tblShift
 *   3. Keep only rows that are:
 *        - on/after retention cutoff (1st of same month, 12 months ago)
 *        - AND date NOT in the CSV date set  (CSV dates are replaced wholesale)
 *   4. Append all CSV rows (no duplicate key checks)
 *   5. Single clear + chunked setValues
 * ============================================================
 */

var EMP_SHIFT_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var EMP_SHIFT_CSV_NAME = 'Employee Shift.csv';
var TBL_SHIFT_HEADERS = ['Emp ID', 'Date', 'Shift'];
var SHIFT_WRITE_CHUNK = 20000;

/** Main entry (Admin Hub / triggers). */
function processShiftFiles() {
  return syncShiftsFromCsv();
}

function syncShiftsFromCsv() {
  var started = new Date().getTime();
  var tz = Session.getScriptTimeZone() || 'Africa/Lagos';
  var cutoff = shiftRetentionCutoff_(); // keep from this date inclusive
  var cutoffTime = cutoff.getTime();

  // ---- 1. CSV: build row list + date set (no object-per-row merge map) ----
  var csvResult = loadEmployeeShiftCsvFast_();
  if (!csvResult.success) {
    return { success: false, message: csvResult.message, kept: 0, purged: 0, csvRows: 0 };
  }
  var csvRows = csvResult.rows;       // [[empId, dateStr, shift], ...]
  var datesInCsv = csvResult.dates;   // { 'yyyy-MM-dd': true }

  Logger.log('CSV loaded: ' + csvRows.length + ' rows, ' +
    Object.keys(datesInCsv).length + ' distinct dates. Elapsed ' +
    (new Date().getTime() - started) + ' ms');

  // ---- 2. Existing tblShift ----
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  if (!sheet) sheet = ss.insertSheet('tblShift');

  var existingData = sheet.getDataRange().getValues();
  Logger.log('tblShift loaded: ' + Math.max(0, existingData.length - 1) +
    ' data rows. Elapsed ' + (new Date().getTime() - started) + ' ms');

  // ---- 3. Keep rows: within retention AND date not covered by CSV ----
  var kept = [];
  var purgedOld = 0;
  var droppedForCsv = 0;

  if (existingData.length > 1) {
    var headers = existingData[0].map(function (h) { return String(h).trim(); });
    var empIdx = headers.indexOf('Emp ID');
    var dateIdx = headers.indexOf('Date');
    var shiftIdx = headers.indexOf('Shift');
    if (empIdx < 0) empIdx = 0;
    if (dateIdx < 0) dateIdx = 1;
    if (shiftIdx < 0) shiftIdx = 2;

    for (var i = 1; i < existingData.length; i++) {
      var row = existingData[i];
      var empId = String(row[empIdx] || '').trim().toUpperCase();
      if (!empId) continue;

      var d = parseShiftDate_(row[dateIdx]);
      if (!d) continue;

      if (d.getTime() < cutoffTime) {
        purgedOld++;
        continue;
      }

      var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (datesInCsv[dateStr]) {
        // Entire calendar day is replaced by CSV — drop old row
        droppedForCsv++;
        continue;
      }

      var shift = String(row[shiftIdx] || '').trim().toUpperCase();
      kept.push([empId, dateStr, shift]);
    }
  }

  Logger.log('Filter done. Kept ' + kept.length + ', purgedOld ' + purgedOld +
    ', droppedForCsv ' + droppedForCsv + '. Elapsed ' +
    (new Date().getTime() - started) + ' ms');

  // ---- 4. Final rows = header + kept + all CSV (no dedupe) ----
  var outRows = [TBL_SHIFT_HEADERS.slice()];
  for (var k = 0; k < kept.length; k++) outRows.push(kept[k]);
  for (var c = 0; c < csvRows.length; c++) outRows.push(csvRows[c]);

  // ---- 5. Write (chunked) ----
  sheet.clearContents();
  writeShiftRowsChunked_(sheet, outRows);
  SpreadsheetApp.flush();

  try {
    if (typeof cacheClearAll_ === 'function') cacheClearAll_();
    CacheService.getScriptCache().remove('emp_map');
  } catch (e) {}

  var ms = new Date().getTime() - started;
  var total = outRows.length - 1;
  var msg = 'Shift sync OK in ' + ms + ' ms. Written ' + total +
    ' rows (kept ' + kept.length + ' + CSV ' + csvRows.length +
    '). Purged old ' + purgedOld + ', replaced CSV dates ' + droppedForCsv +
    '. Cutoff ' + Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd') + '.';
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    total: total,
    kept: kept.length,
    csvRows: csvRows.length,
    purgedOld: purgedOld,
    droppedForCsv: droppedForCsv,
    cutoff: Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd'),
    elapsedMs: ms
  };
}

/**
 * Keep from the 1st of the current month, 12 months ago.
 * Dec 2026 → 2025-12-01 (delete through 2025-11-30).
 * Aug 2026 → 2025-08-01 (delete through 2025-07-31).
 */
function shiftRetentionCutoff_() {
  var now = new Date();
  return new Date(now.getFullYear() - 1, now.getMonth(), 1);
}

/**
 * Load CSV into plain row arrays + date membership set.
 * Filters out rows older than retention cutoff so we never write them back.
 */
function loadEmployeeShiftCsvFast_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SHIFT_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SHIFT_CSV_NAME);
    if (!files.hasNext()) {
      return {
        success: false,
        message: 'CSV not found: "' + EMP_SHIFT_CSV_NAME + '" in folder ' + EMP_SHIFT_FOLDER_ID,
        rows: [],
        dates: {}
      };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var text = file.getBlob().getDataAsString();
    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'Shift CSV is empty.', rows: [], dates: {} };
    }

    var rawHeaders = parsed[0].map(function (h) { return String(h || '').trim(); });
    var headers = rawHeaders.map(function (h) {
      var m = h.match(/^Shift\[(.+)\]$/i);
      return m ? m[1].trim() : h;
    });

    var idx = function (names) {
      for (var n = 0; n < names.length; n++) {
        var want = names[n].toLowerCase();
        for (var k = 0; k < headers.length; k++) {
          if (String(headers[k]).toLowerCase() === want) return k;
        }
      }
      return -1;
    };

    var iEmp = idx(['Emp No', 'Emp ID', 'Employee ID']);
    var iDate = idx(['Tr Date', 'Date', 'Shift Date', 'Transaction Date']);
    var iShift = idx(['Final Shift', 'Shift', 'Shift Code', 'Code']);
    if (iEmp < 0 || iDate < 0 || iShift < 0) {
      return {
        success: false,
        message: 'Shift CSV missing Emp No / Tr Date / Final Shift. Found: ' + headers.join(', '),
        rows: [],
        dates: {}
      };
    }

    var tz = Session.getScriptTimeZone() || 'Africa/Lagos';
    var cutoffTime = shiftRetentionCutoff_().getTime();
    var rows = [];
    var dates = {};

    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row || !row.length) continue;
      var empId = String(row[iEmp] || '').trim().toUpperCase();
      if (!empId) continue;

      var d = parseShiftDate_(row[iDate]);
      if (!d) continue;
      if (d.getTime() < cutoffTime) continue; // never import beyond retention

      var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var shift = String(row[iShift] || '').trim().toUpperCase();
      if (!shift) continue;

      dates[dateStr] = true;
      rows.push([empId, dateStr, shift]);
    }

    return { success: true, rows: rows, dates: dates, count: rows.length };
  } catch (err) {
    return { success: false, message: 'Shift CSV load error: ' + err.message, rows: [], dates: {} };
  }
}

function parseShiftDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function writeShiftRowsChunked_(sheet, rows) {
  if (!rows || !rows.length) return;
  var cols = rows[0].length;
  var chunk = SHIFT_WRITE_CHUNK;
  for (var i = 0; i < rows.length; i += chunk) {
    var part = rows.slice(i, i + chunk);
    sheet.getRange(i + 1, 1, part.length, cols).setValues(part);
    // Yield to the spreadsheet service between large chunks
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
  Logger.log('Daily trigger set: processShiftFiles() ~1:00 AM.');
}
