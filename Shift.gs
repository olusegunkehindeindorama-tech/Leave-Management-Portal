/**
 * ============================================================
 *  SHIFT SYNC — Google Apps Script (trigger-friendly)
 * ============================================================
 *  Source: "Employee Shift.csv" in folder EMP_SHIFT_FOLDER_ID
 *  Target: sheet tblShift  [Emp ID | Date | Shift]
 *
 *  Flow (all in memory → one write):
 *   1. Load CSV
 *   2. Load existing tblShift
 *   3. Drop rows older than the retention cutoff
 *      Cutoff = last day of (current month − 13 months)
 *      e.g. today in Dec 2026 → delete through 30 Nov 2025 (keep from 1 Dec 2025)
 *      e.g. today in Aug 2026 → delete through 31 Jul 2025 (keep from 1 Aug 2025)
 *   4. Merge CSV rows; key = EmpID|yyyy-MM-dd; CSV wins on conflict
 *   5. Single clear + setValues back to tblShift
 * ============================================================
 */

var EMP_SHIFT_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var EMP_SHIFT_CSV_NAME = 'Employee Shift.csv';
var TBL_SHIFT_HEADERS = ['Emp ID', 'Date', 'Shift'];

/**
 * Main entry for manual run and time-driven triggers.
 * Alias kept: processShiftFiles (Admin Hub / old triggers).
 */
function processShiftFiles() {
  return syncShiftsFromCsv();
}

function syncShiftsFromCsv() {
  var started = new Date().getTime();
  var tz = Session.getScriptTimeZone() || 'Africa/Lagos';

  // ---- 1. CSV ----
  var csvResult = loadEmployeeShiftCsv_();
  if (!csvResult.success) {
    return { success: false, message: csvResult.message, kept: 0, added: 0, updated: 0, purged: 0 };
  }
  var csvMap = csvResult.map; // key "EMPID|yyyy-MM-dd" -> { empId, dateStr, shift }

  // ---- 2. Existing tblShift ----
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblShift');
  if (!sheet) {
    sheet = ss.insertSheet('tblShift');
  }

  var existingData = sheet.getDataRange().getValues();
  var existingMap = {}; // same key format
  var purged = 0;
  var keptFromSheet = 0;

  var cutoff = shiftRetentionCutoff_(); // first date to KEEP (inclusive)
  var cutoffTime = cutoff.getTime();

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

      // Drop over-retention window
      if (d.getTime() < cutoffTime) {
        purged++;
        continue;
      }

      var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var key = empId + '|' + dateStr;
      existingMap[key] = {
        empId: empId,
        dateStr: dateStr,
        shift: String(row[shiftIdx] || '').trim().toUpperCase()
      };
      keptFromSheet++;
    }
  }

  // ---- 3. Merge: CSV takes precedence ----
  var added = 0;
  var updated = 0;
  var csvKeys = Object.keys(csvMap);
  for (var c = 0; c < csvKeys.length; c++) {
    var k = csvKeys[c];
    var rec = csvMap[k];

    // Skip CSV rows older than retention (defensive)
    var cd = parseShiftDate_(rec.dateStr);
    if (!cd || cd.getTime() < cutoffTime) continue;

    if (existingMap[k]) {
      if (existingMap[k].shift !== rec.shift) updated++;
    } else {
      added++;
    }
    existingMap[k] = rec; // CSV wins
  }

  // ---- 4. Build output rows (sorted by Emp ID then Date for readability) ----
  var keys = Object.keys(existingMap);
  keys.sort();
  var outRows = [TBL_SHIFT_HEADERS.slice()];
  for (var j = 0; j < keys.length; j++) {
    var r = existingMap[keys[j]];
    outRows.push([r.empId, r.dateStr, r.shift]);
  }

  // ---- 5. Single write ----
  sheet.clearContents();
  writeShiftRowsChunked_(sheet, outRows);

  // Invalidate per-emp shift caches used by leave calc / calendar
  try {
    if (typeof cacheClearAll_ === 'function') cacheClearAll_();
    CacheService.getScriptCache().remove('emp_map');
  } catch (e) {}

  var ms = new Date().getTime() - started;
  var total = keys.length;
  return {
    success: true,
    message: 'Shift sync complete in ' + ms + ' ms. Rows written: ' + total +
      '. Purged (>' + Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd') + ' window start): ' + purged +
      '. CSV applied: ' + csvKeys.length + ' (new keys ~' + added + ', overrides ~' + updated + ').',
    total: total,
    purged: purged,
    added: added,
    updated: updated,
    csvRows: csvKeys.length,
    cutoff: Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd'),
    elapsedMs: ms
  };
}

/**
 * Retention: keep from the 1st of the same calendar month, 12 months ago.
 * Today in Dec 2026 → keep from 2025-12-01 (delete through 2025-11-30).
 * Today in Aug 2026 → keep from 2025-08-01 (delete through 2025-07-31).
 */
function shiftRetentionCutoff_() {
  var now = new Date();
  // 1st of current month, then go back 12 months
  return new Date(now.getFullYear() - 1, now.getMonth(), 1);
}

// ---------------------------------------------------------------------------
// CSV loader
// ---------------------------------------------------------------------------

function loadEmployeeShiftCsv_() {
  try {
    var folder = DriveApp.getFolderById(EMP_SHIFT_FOLDER_ID);
    var files = folder.getFilesByName(EMP_SHIFT_CSV_NAME);
    if (!files.hasNext()) {
      return {
        success: false,
        message: 'CSV not found: "' + EMP_SHIFT_CSV_NAME + '" in folder ' + EMP_SHIFT_FOLDER_ID,
        map: {}
      };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var text = file.getBlob().getDataAsString();
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) {
      return { success: false, message: 'Shift CSV is empty or has no data rows.', map: {} };
    }

    var rawHeaders = rows[0].map(function (h) { return String(h || '').trim(); });
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

    var iEmp = idx(['Emp No', 'Emp ID', 'Employee ID', 'Emp No.']);
    var iDate = idx(['Tr Date', 'Date', 'Shift Date', 'Transaction Date']);
    var iShift = idx(['Final Shift', 'Shift', 'Shift Code', 'Code']);

    if (iEmp < 0 || iDate < 0 || iShift < 0) {
      return {
        success: false,
        message: 'Shift CSV missing required columns (Emp No / Tr Date / Final Shift). Found: ' + headers.join(', '),
        map: {}
      };
    }

    var tz = Session.getScriptTimeZone() || 'Africa/Lagos';
    var map = {};
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var empId = String(row[iEmp] || '').trim().toUpperCase();
      if (!empId) continue;

      var d = parseShiftDate_(row[iDate]);
      if (!d) continue;

      var dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var shift = String(row[iShift] || '').trim().toUpperCase();
      if (!shift) continue;

      var key = empId + '|' + dateStr;
      map[key] = { empId: empId, dateStr: dateStr, shift: shift };
    }

    return { success: true, map: map, count: Object.keys(map).length };
  } catch (err) {
    return { success: false, message: 'Shift CSV load error: ' + err.message, map: {} };
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

/** Write in chunks to stay under Apps Script cell limits for large CSVs. */
function writeShiftRowsChunked_(sheet, rows) {
  if (!rows || !rows.length) return;
  var CHUNK = 10000;
  var cols = rows[0].length;
  for (var i = 0; i < rows.length; i += CHUNK) {
    var chunk = rows.slice(i, i + CHUNK);
    sheet.getRange(i + 1, 1, chunk.length, cols).setValues(chunk);
  }
}

/** Optional: install a daily ~1 AM trigger for syncShiftsFromCsv / processShiftFiles. */
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
