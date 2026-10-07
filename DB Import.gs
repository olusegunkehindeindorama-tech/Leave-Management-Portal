/**
 * DARWINBOX / DB LEAVE IMPORT → daily tblLeave
 * Leave_Application.csv still has From/To dates → expanded to daily rows (noon).
 * Only Approved rows. Dedup: Emp ID | Leave Date.
 */
var LEAVE_CSV_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
var DARWINBOX_CSV_NAME = 'Leave_Application.csv';

function importDarwinBoxLeaves() {
  return importDarwinBoxLeaves_();
}

function importDarwinBoxLeaves_() {
  var started = new Date().getTime();
  Logger.log('=== importDarwinBoxLeaves (daily) START ===');
  if (typeof importRangeLeavesAsDaily_ !== 'function') {
    return { success: false, message: 'importRangeLeavesAsDaily_ missing — load DailyImportBridge.gs' };
  }
  ensureDailyTblLeaveSchema_();
  var parsed = loadDarwinBoxCsvDaily_();
  if (!parsed.success) return parsed;
  var rangeLeaves = [];
  var skippedStatus = 0;
  var bad = 0;
  for (var i = 0; i < parsed.rows.length; i++) {
    var row = parsed.rows[i];
    if (row.status && String(row.status).trim() !== 'Approved') {
      skippedStatus++;
      continue;
    }
    var empId = String(row.empId || '').trim().toUpperCase();
    if (!empId || !row.startDate) { bad++; continue; }
    rangeLeaves.push({
      entryCode: row.entryCode || ('DB-' + empId + '-' + (i + 1)),
      empId: empId,
      leaveType: row.leaveType || 'Annual Leave',
      startDate: row.startDate,
      endDate: row.endDate || row.startDate,
      leaveReason: row.leaveReason || ''
    });
  }
  var result = importRangeLeavesAsDaily_(rangeLeaves, 'DarwinBox');
  result.elapsedMs = new Date().getTime() - started;
  result.skippedStatus = skippedStatus;
  result.badRows = bad;
  result.message = (result.message || '') +
    ' | skipped non-Approved ' + skippedStatus + ' | bad ' + bad +
    ' | ' + result.elapsedMs + ' ms';
  Logger.log('=== importDarwinBoxLeaves END === ' + result.message);
  try { if (typeof scheduleLeaveCleanupPipeline_ === 'function') scheduleLeaveCleanupPipeline_(); } catch (e) {}
  return result;
}

function loadDarwinBoxCsvDaily_() {
  try {
    var folder = DriveApp.getFolderById(LEAVE_CSV_FOLDER_ID);
    var files = folder.getFilesByName(DARWINBOX_CSV_NAME);
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
        if (/leave_application|leave application/i.test(f.getName()) && /\.csv$/i.test(f.getName())) {
          if (!file || f.getLastUpdated() > file.getLastUpdated()) file = f;
        }
      }
    }
    if (!file) return { success: false, message: 'CSV not found: ' + DARWINBOX_CSV_NAME };
    var text = file.getBlob().getDataAsString().replace(/^\uFEFF/, '');
    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) return { success: false, message: 'DarwinBox CSV empty' };
    var headers = parsed[0].map(function (h) { return String(h || '').trim(); });
    function find(names) {
      for (var n = 0; n < names.length; n++) {
        var w = names[n].toLowerCase();
        for (var k = 0; k < headers.length; k++) {
          if (String(headers[k]).toLowerCase() === w) return k;
        }
      }
      return -1;
    }
    var iEmp = find(['Employee Id', 'Employee ID', 'Emp ID']);
    var iStart = find(['Leave From Date', 'From Date', 'Start Date']);
    var iEnd = find(['Leave To Date', 'To Date', 'End Date']);
    var iStatus = find(['Status']);
    var iType = find(['Leave Type']);
    var iComment = find(['Employee Comment', 'Comment', 'Reason', 'Leave Reason']);
    var iEntry = find(['Entry Code', 'Request Id', 'Request ID']);
    if (iEmp < 0 || iStart < 0 || iEnd < 0) {
      return { success: false, message: 'DarwinBox CSV missing columns. Found: ' + headers.join(', ') };
    }
    var out = [];
    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row || !row.length) continue;
      out.push({
        entryCode: iEntry >= 0 ? String(row[iEntry] || '').trim() : '',
        empId: String(row[iEmp] || '').trim(),
        leaveType: iType >= 0 ? String(row[iType] || '').trim() : 'Annual Leave',
        startDate: String(row[iStart] || '').trim(),
        endDate: String(row[iEnd] || '').trim(),
        leaveReason: iComment >= 0 ? String(row[iComment] || '').trim() : '',
        status: iStatus >= 0 ? String(row[iStatus] || '').trim() : 'Approved'
      });
    }
    return { success: true, rows: out, headers: headers };
  } catch (err) {
    return { success: false, message: 'DarwinBox CSV error: ' + err.message };
  }
}
