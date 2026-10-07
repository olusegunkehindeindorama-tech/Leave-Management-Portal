/**
 * EXCEL LEAVE IMPORT → daily tblLeave
 * Source still has Start/End ranges → expanded to one row per day (noon).
 * Dedup: Emp ID | Leave Date.
 */
var EXCEL_LEAVE_CSV_NAME = 'Excel Leave Entries.csv';
var EXCEL_LEAVE_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';

function importExcelLeaves() { return importExcelLeaves_(); }
function importExcelLeaveEntries() { return importExcelLeaves_(); }

function importExcelLeaves_() {
  var started = new Date().getTime();
  Logger.log('=== importExcelLeaves (daily) START ===');
  if (typeof importRangeLeavesAsDaily_ !== 'function') {
    return { success: false, message: 'importRangeLeavesAsDaily_ missing — load DailyImportBridge.gs' };
  }
  ensureDailyTblLeaveSchema_();
  var parsed = loadExcelLeaveCsvDaily_();
  if (!parsed.success) return parsed;
  var rangeLeaves = [];
  var bad = 0;
  for (var i = 0; i < parsed.rows.length; i++) {
    var row = parsed.rows[i];
    var empId = String(row.empId || '').trim().toUpperCase();
    if (!empId || !row.startDate) { bad++; continue; }
    rangeLeaves.push({
      entryCode: row.entryCode || ('XL-' + empId + '-' + (i + 1)),
      empId: empId,
      leaveType: row.leaveType || 'Annual Leave',
      startDate: row.startDate,
      endDate: row.endDate || row.startDate,
      leaveReason: row.leaveReason || ''
    });
  }
  var result = importRangeLeavesAsDaily_(rangeLeaves, 'Excel');
  result.elapsedMs = new Date().getTime() - started;
  result.badRows = bad;
  result.message = (result.message || '') + ' in ' + result.elapsedMs + ' ms';
  Logger.log('=== importExcelLeaves END === ' + result.message);
  try { if (typeof scheduleLeaveCleanupPipeline_ === 'function') scheduleLeaveCleanupPipeline_(); } catch (e) {}
  return result;
}

function loadExcelLeaveCsvDaily_() {
  try {
    var folder = DriveApp.getFolderById(EXCEL_LEAVE_FOLDER_ID);
    var files = folder.getFilesByName(EXCEL_LEAVE_CSV_NAME);
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
        if (/excel.*leave.*\.csv$/i.test(f.getName()) || /leave\s*entr/i.test(f.getName())) {
          if (!file || f.getLastUpdated() > file.getLastUpdated()) file = f;
        }
      }
    }
    if (!file) return { success: false, message: 'CSV not found: ' + EXCEL_LEAVE_CSV_NAME };
    var text = file.getBlob().getDataAsString().replace(/^\uFEFF/, '');
    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) return { success: false, message: 'Excel leave CSV empty' };
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
    var iEntry = find(['Entry Code', 'EntryCode', 'Leave Code']);
    var iEmp = find(['Emp ID', 'Emp No', 'Employee ID', 'Employee Id']);
    var iType = find(['Leave Type', 'Type']);
    var iStart = find(['Start Date', 'From Date', 'Leave From Date', 'Start']);
    var iEnd = find(['End Date', 'To Date', 'Leave To Date', 'End']);
    var iReason = find(['Leave Reason', 'Reason', 'Comment']);
    if (iEmp < 0 || iStart < 0) {
      return { success: false, message: 'Excel CSV missing Emp ID / Start Date. Found: ' + headers.join(', ') };
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
        endDate: iEnd >= 0 ? String(row[iEnd] || '').trim() : String(row[iStart] || '').trim(),
        leaveReason: iReason >= 0 ? String(row[iReason] || '').trim() : ''
      });
    }
    return { success: true, rows: out, headers: headers };
  } catch (err) {
    return { success: false, message: 'Excel CSV error: ' + err.message };
  }
}
