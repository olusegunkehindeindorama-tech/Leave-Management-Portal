/**
 * ============================================================
 *  EXCEL LEAVE ENTRIES IMPORT — on-click only (not for triggers)
 * ============================================================
 *  Source: "Excel Leave Entries.csv" in folder LEAVE_CSV_FOLDER_ID
 *  Target: tblLeave — append rows not already present
 *
 *  CSV columns (clean export):
 *    Leave Code, Emp ID, Emp Name, Department, Category, Leave Type,
 *    Start Date, End Date, Leave Reason, Date Entered, Entered By,
 *    Date Modified, Modified By, BU, DB Remark, Upload Date,
 *    Uploaded By, DB Leave Code
 *
 *  Mapping → tblLeave:
 *    Leave Code     → Entry Code
 *    DB Leave Code  → Leave Code
 *    Dates: dd-MMM-yy (31-Mar-26), dd/MM/yyyy HH:mm (26/05/2026 00:00)
 *
 *  Dedup: Entry Code (if present) OR EmpID|Start|End
 *  No of Days / Leave Utilized / Entitlement Year left blank
 * ============================================================
 */

var EXCEL_LEAVE_CSV_NAME = 'Excel Leave Entries.csv';

/** Manual / on-click entry point. */
function importExcelLeaveEntries() {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName('tblLeave');
  if (!leaveSheet) {
    return { success: false, message: 'tblLeave sheet missing.' };
  }

  var ctx = loadLeaveImportContext_(leaveSheet);
  if (!ctx.success) return ctx;

  var parsed = loadExcelLeaveEntriesCsv_();
  if (!parsed.success) return parsed;

  var headers = parsed.headers;
  var rows = parsed.rows;

  var idx = {
    entryCode: findHeader_(headers, ['Leave Code', 'Entry Code']),
    dbLeaveCode: findHeader_(headers, ['DB Leave Code']),
    empId: findHeader_(headers, ['Emp ID', 'Employee Id', 'Employee ID']),
    empName: findHeader_(headers, ['Emp Name', 'Employee Name']),
    dept: findHeader_(headers, ['Department', 'Dept']),
    category: findHeader_(headers, ['Category']),
    leaveType: findHeader_(headers, ['Leave Type']),
    start: findHeader_(headers, ['Start Date', 'Leave From Date']),
    end: findHeader_(headers, ['End Date', 'Leave To Date']),
    reason: findHeader_(headers, ['Leave Reason', 'Reason', 'Employee Comment']),
    dateEntered: findHeader_(headers, ['Date Entered', 'Applied On']),
    enteredBy: findHeader_(headers, ['Entered By']),
    dateModified: findHeader_(headers, ['Date Modified']),
    modifiedBy: findHeader_(headers, ['Modified By']),
    bu: findHeader_(headers, ['BU', 'Business Unit']),
    dbRemark: findHeader_(headers, ['DB Remark', 'Remark']),
    uploadDate: findHeader_(headers, ['Upload Date']),
    uploadedBy: findHeader_(headers, ['Uploaded By', 'Upload By'])
  };

  // Ensure DB Leave Code is not confused with Leave Code (Entry Code)
  if (idx.entryCode >= 0 && idx.dbLeaveCode === idx.entryCode) {
    idx.dbLeaveCode = -1;
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim().toLowerCase() === 'db leave code') {
        idx.dbLeaveCode = h;
        break;
      }
    }
  }

  if (idx.empId < 0 || idx.start < 0 || idx.end < 0) {
    return {
      success: false,
      message: 'Excel Leave CSV missing Emp ID / Start / End. Found: ' + headers.join(', ')
    };
  }

  Logger.log('Excel CSV columns mapped — Entry:' + idx.entryCode +
    ' DBCode:' + idx.dbLeaveCode + ' Emp:' + idx.empId +
    ' Start:' + idx.start + ' End:' + idx.end +
    ' rows:' + rows.length);

  var newRows = [];
  var skippedDup = 0;
  var skippedBad = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row || !row.length) continue;

    var empId = String(row[idx.empId] || '').trim().toUpperCase();
    var startDate = parseLeaveDate_(row[idx.start]);
    var endDate = parseLeaveDate_(row[idx.end]);
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
    var dbCode = idx.dbLeaveCode >= 0 ? String(row[idx.dbLeaveCode] || '').trim() : '';

    if (leaveType && ctx.policyMap[leaveType.toLowerCase()]) {
      var pol = ctx.policyMap[leaveType.toLowerCase()];
      if (!dbCode) dbCode = pol.dbCode;
    }

    if (!entryCode) entryCode = 'XL-' + Date.now() + '-' + i;

    var dept = idx.dept >= 0 ? String(row[idx.dept] || '').trim() : '';
    var cat = idx.category >= 0 ? String(row[idx.category] || '').trim() : '';
    var bu = idx.bu >= 0 ? String(row[idx.bu] || '').trim() : '';
    var empName = idx.empName >= 0 ? String(row[idx.empName] || '').trim() : '';

    if (!dept) dept = empInfo.dept;
    if (!cat) cat = empInfo.cat;
    if (!bu) bu = empInfo.bu;
    if (!empName) empName = empInfo.name;

    var newRow = buildBlankLeaveRow_(ctx.lHeaders);
    setLeaveCol_(newRow, ctx.lHeaders, 'Entry Code', entryCode);
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Code', dbCode);
    setLeaveCol_(newRow, ctx.lHeaders, 'Emp ID', empId);
    setLeaveCol_(newRow, ctx.lHeaders, 'Emp Name', empName);
    setLeaveCol_(newRow, ctx.lHeaders, 'Department', dept);
    setLeaveCol_(newRow, ctx.lHeaders, 'Category', cat);
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Type', leaveType);
    setLeaveCol_(newRow, ctx.lHeaders, 'Start Date', startDate);
    setLeaveCol_(newRow, ctx.lHeaders, 'End Date', endDate);
    setLeaveCol_(newRow, ctx.lHeaders, 'Leave Reason',
      idx.reason >= 0 ? String(row[idx.reason] || '').trim() : '');
    setLeaveCol_(newRow, ctx.lHeaders, 'Date Entered',
      idx.dateEntered >= 0 ? (parseLeaveDate_(row[idx.dateEntered]) || '') : '');
    setLeaveCol_(newRow, ctx.lHeaders, 'Entered By',
      idx.enteredBy >= 0 ? String(row[idx.enteredBy] || '').trim() : 'Excel Import');
    setLeaveCol_(newRow, ctx.lHeaders, 'Date Modified',
      idx.dateModified >= 0 ? (parseLeaveDate_(row[idx.dateModified]) || '') : '');
    setLeaveCol_(newRow, ctx.lHeaders, 'Modified By',
      idx.modifiedBy >= 0 ? String(row[idx.modifiedBy] || '').trim() : '');
    setLeaveCol_(newRow, ctx.lHeaders, 'BU', bu);
    setLeaveCol_(newRow, ctx.lHeaders, 'DB Remark',
      idx.dbRemark >= 0 ? String(row[idx.dbRemark] || '').trim() : 'Excel Import');
    setLeaveCol_(newRow, ctx.lHeaders, 'Upload Date',
      idx.uploadDate >= 0 ? (parseLeaveDate_(row[idx.uploadDate]) || new Date()) : new Date());
    setLeaveCol_(newRow, ctx.lHeaders, 'Uploaded By',
      idx.uploadedBy >= 0 ? String(row[idx.uploadedBy] || '').trim() : 'Excel Import');

    newRows.push(newRow);
    ctx.existingKeys[fp] = true;
    if (entryKey) ctx.existingEntryCodes[entryKey] = true;
  }

  if (newRows.length) {
    appendLeaveRows_(leaveSheet, newRows);
  }

  var ms = new Date().getTime() - started;
  var msg = 'Excel Leave import: +' + newRows.length + ' new (skipped dup ' +
    skippedDup + ', bad ' + skippedBad + ') from ' + rows.length +
    ' CSV rows in ' + ms + ' ms.';
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    added: newRows.length,
    skippedDup: skippedDup,
    skippedBad: skippedBad,
    csvRows: rows.length,
    elapsedMs: ms
  };
}

/**
 * Load clean Excel Leave Entries.csv (standard CSV).
 * Falls back to unwrapping a legacy broken export envelope if needed.
 */
function loadExcelLeaveEntriesCsv_() {
  try {
    var folder = DriveApp.getFolderById(LEAVE_CSV_FOLDER_ID);
    var files = folder.getFilesByName(EXCEL_LEAVE_CSV_NAME);
    if (!files.hasNext()) {
      return { success: false, message: 'CSV not found: "' + EXCEL_LEAVE_CSV_NAME + '"' };
    }
    var file = files.next();
    while (files.hasNext()) {
      var f2 = files.next();
      if (f2.getLastUpdated() > file.getLastUpdated()) file = f2;
    }

    var text = file.getBlob().getDataAsString();

    // Normal clean CSV starts with header
    var trimmed = text.replace(/^\uFEFF/, '').trim();
    if (trimmed.indexOf('Leave Code,Emp ID') !== 0 &&
        trimmed.indexOf('Leave Code, Emp ID') !== 0) {
      // Legacy broken wrapper fallback
      var headerPos = text.indexOf('Leave Code,Emp ID');
      if (headerPos < 0) headerPos = text.indexOf('Leave Code, Emp ID');
      if (headerPos >= 0) {
        var body = text.substring(headerPos);
        body = body.replace(/\\r\\n/g, '\n').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        body = body.replace(/"?\s*,?\s*null\s*,?\s*null\s*,?\s*null\s*\]?\s*$/i, '');
        body = body.replace(/"\s*$/, '');
        text = body;
      }
    }

    var parsed = Utilities.parseCsv(text);
    if (!parsed || parsed.length < 2) {
      return { success: false, message: 'Excel Leave CSV empty or unreadable.' };
    }

    var cleaned = [];
    cleaned.push(parsed[0].map(function (h) { return String(h || '').trim(); }));
    for (var i = 1; i < parsed.length; i++) {
      var r = parsed[i];
      if (!r || !r.length) continue;
      var first = String(r[0] || '').trim();
      if (!first || first.toLowerCase() === 'null') continue;
      cleaned.push(r);
    }

    if (cleaned.length < 2) {
      return { success: false, message: 'No data rows found in Excel Leave CSV.' };
    }

    return {
      success: true,
      headers: cleaned[0],
      rows: cleaned.slice(1),
      count: cleaned.length - 1
    };
  } catch (err) {
    return { success: false, message: 'Excel Leave CSV error: ' + err.message };
  }
}
