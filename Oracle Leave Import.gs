/**
 * ORACLE LEAVE IMPORT → daily tblLeave rows
 * CSV still has Start/End ranges → expanded to one row per day (noon dates).
 * Dedup: Emp ID | Leave Date.
 * Requires: DailyLeaveModel.gs, DailyImportBridge.gs
 */
var ORACLE_LEAVE_CSV_NAME = 'Oracle Leave.csv';
var ORACLE_LEAVE_FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';

function importOracleLeaves() {
  return importOracleLeaves_();
}

function importOracleLeaves_() {
  var started = new Date().getTime();
  Logger.log('=== importOracleLeaves (daily) START ===');

  if (typeof importRangeLeavesAsDaily_ !== 'function') {
    return { success: false, message: 'importRangeLeavesAsDaily_ missing — load DailyImportBridge.gs' };
  }
  if (typeof ensureDailyTblLeaveSchema_ !== 'function') {
    return { success: false, message: 'DailyLeaveModel.gs not loaded' };
  }

  ensureDailyTblLeaveSchema_();

  var parsed = loadOracleLeaveCsvDaily_();
  if (!parsed.success) return parsed;

  var rangeLeaves = [];
  var bad = 0;
  for (var i = 0; i < parsed.rows.length; i++) {
    var row = parsed.rows[i];
    var empId = String(row.empId || '').trim().toUpperCase();
    if (!empId || !row.startDate) { bad++; continue; }
    rangeLeaves.push({
      entryCode: row.entryCode || ('OR-' + empId + '-' + (i + 1)),
      empId: empId,
      leaveType: row.leaveType || 'Annual Leave',
      startDate: row.startDate,
      endDate: row.endDate || row.startDate,
      leaveReason: row.leaveReason || ''
    });
  }

  var result = importRangeLeavesAsDaily_(rangeLeaves, 'Oracle');
  var ms = new Date().getTime() - started;
  result.elapsedMs = ms;
  result.badRows = bad;
  result.message = (result.message || '') + ' in ' + ms + ' ms (bad source rows: ' + bad + ')';
  Logger.log('=== importOracleLeaves END === ' + result.message);

  try {
    if (typeof scheduleLeaveCleanupPipeline_ === 'function') scheduleLeaveCleanupPipeline_();
  } catch (e) {}
  return result;
}

function loadOracleLeaveCsvDaily_() {
  try {
    var folder = DriveApp.getFolderById(ORACLE_LEAVE_FOLDER_ID);
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
        if (/oracle\s*leave.*\.csv$/i.test(f.getName())) {
          if (!file || f.getLastUpdated() > file.getLastUpdated()) file = f;
        }
      }
    }
    if (!file) {
      return { success: false, message: 'CSV not found: ' + ORACLE_LEAVE_CSV_NAME };
    }

    var text = file.getBlob().getDataAsString().replace(/^\uFEFF/, '');
    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'Oracle CSV empty' };
    }

    var rawH = parsed[0].map(function (h) { return String(h || '').trim(); });
    var headers = rawH.map(function (h) {
      var m = h.match(/^Oracle Leave\[(.+)\]$/i);
      return m ? m[1].trim() : h.replace(/^Oracle Leave\s*/i, '').trim();
    });

    function find(names) {
      for (var n = 0; n < names.length; n++) {
        var w = names[n].toLowerCase();
        for (var k = 0; k < headers.length; k++) {
          if (String(headers[k]).toLowerCase() === w) return k;
        }
      }
      return -1;
    }

    var iEntry = find(['Entry Code']);
    var iEmp = find(['Emp ID', 'Emp No', 'Employee ID']);
    var iType = find(['Leave Type']);
    var iStart = find(['Start Date', 'From Date', 'Leave From Date']);
    var iEnd = find(['End Date', 'To Date', 'Leave To Date']);
    var iReason = find(['Leave Reason', 'Reason', 'Comment']);
    if (iEmp < 0 || iStart < 0) {
      return { success: false, message: 'Oracle CSV missing Emp ID / Start Date. Found: ' + headers.join(', ') };
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
    Logger.log('Oracle CSV rows: ' + out.length);
    return { success: true, rows: out, headers: headers };
  } catch (err) {
    return { success: false, message: 'Oracle CSV error: ' + err.message };
  }
}
