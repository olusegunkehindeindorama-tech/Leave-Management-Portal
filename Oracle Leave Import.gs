/**
 * ORACLE LEAVE IMPORT — on-click only
 * Source file: "Oracle Leave.csv" in the Darwinbox download folder (LEAVE_CSV_FOLDER_ID).
 * After append: scheduleLeaveCleanupPipeline_() (~6 min) — not run inline.
 *
 * Column mapping (CSV → tblLeave):
 *   Entry Code / Oracle Leave[Entry Code]   → Entry Code
 *   Leave Code / Oracle Leave[Leave Code]   → Leave Code
 *   Emp ID / Oracle Leave[Emp ID]           → Emp ID
 *   Emp Name / Oracle Leave[Emp Name]       → Emp Name
 *   Department / Oracle Leave[Department]   → Department
 *   Category / Oracle Leave[Category]       → Category
 *   Leave Type / Oracle Leave[Leave Type]   → Leave Type
 *   Start Date / Oracle Leave[Start Date]   → Start Date  (noon local via toSheetDateValue_)
 *   End Date / Oracle Leave[End Date]       → End Date    (noon local)
 *   Leave Reason / Oracle Leave[Leave Reason] → Leave Reason
 *   Date Entered / Oracle Leave[Date Entered] → Date Entered
 *   BU / Oracle Leave[BU]                   → BU
 *
 * Auto-filled (not in CSV):
 *   Entered By   = "Oracle"
 *   DB Remark    = "Pending"
 *   Upload Date  = now
 *   Uploaded By  = "Oracle Import"
 *   Other tblLeave columns left blank (No of Days, Leave Utilized, etc.
 *   are filled later by recalculateAllLeaveUtilized).
 */
var ORACLE_LEAVE_CSV_NAME = 'Oracle Leave.csv';

function importOracleLeaves() {
  return importOracleLeaves_();
}

function importOracleLeaves_() {
  var started = new Date().getTime();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leaveSheet = ss.getSheetByName('tblLeave');
    if (!leaveSheet) {
      return { success: false, message: 'tblLeave sheet missing.' };
    }
    if (typeof loadLeaveImportContext_ !== 'function') {
      return {
        success: false,
        message: 'loadLeaveImportContext_ missing — load DB Import.gs / DB Import Helpers.gs.'
      };
    }
    if (typeof appendLeaveRows_ !== 'function') {
      return { success: false, message: 'appendLeaveRows_ missing — load DB Import Helpers.gs.' };
    }

    var ctx = loadLeaveImportContext_(leaveSheet);
    if (!ctx.success) return ctx;

    var parsed = loadOracleLeaveCsv_();
    if (!parsed.success) return parsed;

    var headers = parsed.headers;
    var rows = parsed.rows;

    var idx = {
      entryCode: oracleFindHeader_(headers, [
        'Entry Code', 'Oracle Leave[Entry Code]', 'Oracle Leave [Entry Code]'
      ]),
      leaveCode: oracleFindHeader_(headers, [
        'Leave Code', 'Oracle Leave[Leave Code]', 'Oracle Leave [Leave Code]', 'DB Leave Code'
      ]),
      empId: oracleFindHeader_(headers, [
        'Emp ID', 'Employee Id', 'Employee ID',
        'Oracle Leave[Emp ID]', 'Oracle Leave [Emp ID]'
      ]),
      empName: oracleFindHeader_(headers, [
        'Emp Name', 'Employee Name',
        'Oracle Leave[Emp Name]', 'Oracle Leave [Emp Name]'
      ]),
      dept: oracleFindHeader_(headers, [
        'Department', 'Dept',
        'Oracle Leave[Department]', 'Oracle Leave [Department]'
      ]),
      category: oracleFindHeader_(headers, [
        'Category',
        'Oracle Leave[Category]', 'Oracle Leave [Category]'
      ]),
      leaveType: oracleFindHeader_(headers, [
        'Leave Type',
        'Oracle Leave[Leave Type]', 'Oracle Leave [Leave Type]'
      ]),
      start: oracleFindHeader_(headers, [
        'Start Date', 'Leave From Date',
        'Oracle Leave[Start Date]', 'Oracle Leave [Start Date]'
      ]),
      end: oracleFindHeader_(headers, [
        'End Date', 'Leave To Date',
        'Oracle Leave[End Date]', 'Oracle Leave [End Date]'
      ]),
      reason: oracleFindHeader_(headers, [
        'Leave Reason', 'Reason',
        'Oracle Leave[Leave Reason]', 'Oracle Leave [Leave Reason]'
      ]),
      dateEntered: oracleFindHeader_(headers, [
        'Date Entered', 'Applied On',
        'Oracle Leave[Date Entered]', 'Oracle Leave [Date Entered]'
      ]),
      bu: oracleFindHeader_(headers, [
        'BU', 'Business Unit',
        'Oracle Leave[BU]', 'Oracle Leave [BU]'
      ])
    };

    if (idx.empId < 0 || idx.start < 0 || idx.end < 0) {
      return {
        success: false,
        message: 'Oracle Leave CSV missing Emp ID / Start / End. Found: ' + headers.join(', ')
      };
    }

    var newRows = [];
    var skippedDup = 0;
    var skippedBad = 0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row || !row.length) continue;

      var empId = String(row[idx.empId] || '').trim().toUpperCase();
      var startDate = oracleResolveDate_(row[idx.start]);
      var endDate = oracleResolveDate_(row[idx.end]);
      if (!empId || !startDate || !endDate) {
        skippedBad++;
        continue;
      }

      var entryCode = idx.entryCode >= 0 ? String(row[idx.entryCode] || '').trim() : '';
      var entryKey = entryCode ? entryCode.toUpperCase() : '';

      if (entryKey && ctx.existingEntryCodes[entryKey]) {
        skippedDup++;
        continue;
      }
      var fp = fingerprintKey_(empId, startDate, endDate);
      if (ctx.existingKeys[fp]) {
        skippedDup++;
        continue;
      }

      var empInfo = ctx.empMap[empId] || { bu: '', cat: '', dept: '', name: '' };
      var leaveType = idx.leaveType >= 0 ? String(row[idx.leaveType] || '').trim() : '';
      var leaveCode = idx.leaveCode >= 0 ? String(row[idx.leaveCode] || '').trim() : '';

      if (leaveType && ctx.policyMap[leaveType.toLowerCase()]) {
        var pol = ctx.policyMap[leaveType.toLowerCase()];
        if (!leaveCode) leaveCode = pol.dbCode || '';
      }

      if (!entryCode) entryCode = 'OR-' + Date.now() + '-' + i;

      var dept = idx.dept >= 0 ? String(row[idx.dept] || '').trim() : '';
      var cat = idx.category >= 0 ? String(row[idx.category] || '').trim() : '';
      var bu = idx.bu >= 0 ? String(row[idx.bu] || '').trim() : '';
      var empName = idx.empName >= 0 ? String(row[idx.empName] || '').trim() : '';
      if (!dept) dept = empInfo.dept || '';
      if (!cat) cat = empInfo.cat || '';
      if (!bu) bu = empInfo.bu || '';
      if (!empName) empName = empInfo.name || '';

      var newRow = buildBlankLeaveRow_(ctx.lHeaders);
      setLeaveCol_(newRow, ctx.lHeaders, 'Entry Code', entryCode);
      setLeaveCol_(newRow, ctx.lHeaders, 'Leave Code', leaveCode);
      setLeaveCol_(newRow, ctx.lHeaders, 'Emp ID', empId);
      setLeaveCol_(newRow, ctx.lHeaders, 'Emp Name', empName);
      setLeaveCol_(newRow, ctx.lHeaders, 'Department', dept);
      setLeaveCol_(newRow, ctx.lHeaders, 'Category', cat);
      setLeaveCol_(newRow, ctx.lHeaders, 'Leave Type', leaveType);
      setLeaveCol_(newRow, ctx.lHeaders, 'Start Date', oracleSheetDate_(startDate));
      setLeaveCol_(newRow, ctx.lHeaders, 'End Date', oracleSheetDate_(endDate));
      setLeaveCol_(newRow, ctx.lHeaders, 'Leave Reason',
        idx.reason >= 0 ? String(row[idx.reason] || '').trim() : '');
      setLeaveCol_(newRow, ctx.lHeaders, 'Date Entered',
        idx.dateEntered >= 0
          ? (oracleResolveDate_(row[idx.dateEntered]) || new Date())
          : new Date());
      setLeaveCol_(newRow, ctx.lHeaders, 'Entered By', 'Oracle');
      setLeaveCol_(newRow, ctx.lHeaders, 'BU', bu);
      setLeaveCol_(newRow, ctx.lHeaders, 'DB Remark', 'Pending');
      setLeaveCol_(newRow, ctx.lHeaders, 'Upload Date', new Date());
      setLeaveCol_(newRow, ctx.lHeaders, 'Uploaded By', 'Oracle Import');

      newRows.push(newRow);
      ctx.existingKeys[fp] = true;
      if (entryKey) ctx.existingEntryCodes[entryKey] = true;
    }

    if (newRows.length) {
      appendLeaveRows_(leaveSheet, newRows);
      try {
        var last = leaveSheet.getLastRow();
        var first = Math.max(2, last - newRows.length + 1);
        var h = leaveSheet.getRange(1, 1, 1, leaveSheet.getLastColumn()).getValues()[0]
          .map(function (x) { return String(x).trim(); });
        var si = h.indexOf('Start Date');
        var ei = h.indexOf('End Date');
        if (si >= 0) leaveSheet.getRange(first, si + 1, newRows.length, 1).setNumberFormat('dd-mmm-yyyy');
        if (ei >= 0) leaveSheet.getRange(first, ei + 1, newRows.length, 1).setNumberFormat('dd-mmm-yyyy');
      } catch (fe) {
        Logger.log('Oracle date format: ' + fe.message);
      }
    }

    var schedule = null;
    if (typeof scheduleLeaveCleanupPipeline_ === 'function') {
      try {
        schedule = scheduleLeaveCleanupPipeline_();
      } catch (se) {
        schedule = { success: false, message: se.message };
        Logger.log('Could not schedule cleanup pipeline: ' + se.message);
      }
    }

    var ms = new Date().getTime() - started;
    var msg = 'Oracle Leave import: +' + newRows.length + ' new (skipped dup ' +
      skippedDup + ', bad ' + skippedBad + ') from ' + rows.length +
      ' CSV rows in ' + ms + ' ms.';
    if (schedule && schedule.message) msg += ' | ' + schedule.message;
    Logger.log(msg);
    return {
      success: true,
      message: msg,
      added: newRows.length,
      skippedDup: skippedDup,
      skippedBad: skippedBad,
      schedule: schedule,
      elapsedMs: ms,
      csvRows: rows.length
    };
  } catch (err) {
    Logger.log('importOracleLeaves_ ERROR: ' + err.message + '\n' + err.stack);
    return {
      success: false,
      message: 'Oracle Leave import failed: ' + err.message,
      error: String(err.stack || err)
    };
  }
}

function oracleFindHeader_(headers, names) {
  if (typeof findHeader_ === 'function') {
    var hit = findHeader_(headers, names);
    if (hit >= 0) return hit;
  }
  var normalized = headers.map(function (h) {
    var s = String(h || '').trim();
    var m = s.match(/^Oracle\s*Leave\s*\[\s*(.+?)\s*\]$/i);
    if (m) s = m[1].trim();
    return s.toLowerCase();
  });
  for (var n = 0; n < names.length; n++) {
    var want = String(names[n] || '').trim().toLowerCase();
    var wm = want.match(/^oracle\s*leave\s*\[\s*(.+?)\s*\]$/i);
    if (wm) want = wm[1].trim();
    for (var i = 0; i < normalized.length; i++) {
      if (normalized[i] === want) return i;
    }
  }
  for (var n2 = 0; n2 < names.length; n2++) {
    var w2 = String(names[n2] || '').trim().toLowerCase()
      .replace(/^oracle\s*leave\s*\[\s*/i, '').replace(/\s*\]$/i, '');
    for (var j = 0; j < normalized.length; j++) {
      if (normalized[j].indexOf(w2) >= 0 || w2.indexOf(normalized[j]) >= 0) return j;
    }
  }
  return -1;
}

function oracleResolveDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof toCalendarDate_ === 'function') {
    var d = toCalendarDate_(val);
    if (d) return d;
  }
  if (typeof parseLeaveDate_ === 'function') {
    var p = parseLeaveDate_(val);
    if (p) return p;
  }
  var s = String(val).trim();
  var dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (dmy) {
    var day = Number(dmy[1]);
    var month = Number(dmy[2]);
    var year = Number(dmy[3]);
    if (year < 100) year = year >= 70 ? 1900 + year : 2000 + year;
    return new Date(year, month - 1, day);
  }
  var dt = new Date(s);
  if (!isNaN(dt.getTime())) {
    return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  return null;
}

function oracleSheetDate_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof toSheetDateValue_ === 'function') {
    var v = toSheetDateValue_(val);
    if (v !== '' && v !== null && v !== undefined) return v;
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate(), 12, 0, 0);
  }
  var d2 = oracleResolveDate_(val);
  if (!d2) return '';
  return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0);
}

function loadOracleLeaveCsv_() {
  try {
    if (typeof LEAVE_CSV_FOLDER_ID === 'undefined' || !LEAVE_CSV_FOLDER_ID) {
      return {
        success: false,
        message: 'LEAVE_CSV_FOLDER_ID not defined — load DB Import.gs first.'
      };
    }
    var folder = DriveApp.getFolderById(LEAVE_CSV_FOLDER_ID);
    var files = folder.getFilesByName(ORACLE_LEAVE_CSV_NAME);
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
        var nm = f.getName();
        if (/^oracle\s*leave\.csv$/i.test(nm) || /^oracle.?leave.*\.csv$/i.test(nm)) {
          if (!file || f.getLastUpdated() > file.getLastUpdated()) file = f;
        }
      }
    }
    if (!file) {
      return { success: false, message: 'CSV not found: "' + ORACLE_LEAVE_CSV_NAME + '"' };
    }

    var text = file.getBlob().getDataAsString();
    text = text.replace(/^\uFEFF/, '');
    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'Oracle Leave CSV empty or unreadable.' };
    }

    var headers = parsed[0].map(function (h) { return String(h || '').trim(); });
    var cleaned = [];
    for (var i = 1; i < parsed.length; i++) {
      var r = parsed[i];
      if (!r || !r.length) continue;
      var hasAny = false;
      for (var c = 0; c < r.length; c++) {
        if (String(r[c] || '').trim()) { hasAny = true; break; }
      }
      if (!hasAny) continue;
      cleaned.push(r);
    }

    if (!cleaned.length) {
      return { success: false, message: 'No data rows found in Oracle Leave CSV.' };
    }

    return {
      success: true,
      headers: headers,
      rows: cleaned,
      count: cleaned.length,
      fileName: file.getName()
    };
  } catch (err) {
    return { success: false, message: 'Oracle Leave CSV error: ' + err.message };
  }
}
