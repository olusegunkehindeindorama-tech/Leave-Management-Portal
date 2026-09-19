/**
 * Leave records export by overlapping date range.
 *
 * Filter: leave.EndDate >= rangeStart AND leave.StartDate <= rangeEnd
 *
 * Columns:
 *   Entry Code, Emp ID, Emp Name, Department, Category, Leave Type,
 *   Start Date, End Date, Leave Reason, No of Days (calendar),
 *   Leave Utilized (work days), Entitlement Year, Date Entered, Entered By
 */

/**
 * @param {string|Date} startDate - range start (inclusive)
 * @param {string|Date} endDate   - range end (inclusive)
 * @returns {{success:boolean, rows:Array, csv:string, message:string}}
 */
function exportLeaveRecordsByDateRange(startDate, endDate) {
  var rangeStart = parseExportDate_(startDate);
  var rangeEnd = parseExportDate_(endDate);
  if (!rangeStart || !rangeEnd) {
    return { success: false, message: 'Provide valid start and end dates.', rows: [], csv: '' };
  }
  if (rangeEnd < rangeStart) {
    return { success: false, message: 'End date must be on or after start date.', rows: [], csv: '' };
  }

  var result = getLeaveRecordsInRange_(rangeStart, rangeEnd);
  var csv = arrayToCsv_(result.table);

  return {
    success: true,
    message: result.rows.length + ' leave record(s) overlapping ' +
      formatExportDate_(rangeStart) + ' to ' + formatExportDate_(rangeEnd),
    rows: result.rows,
    table: result.table,
    csv: csv,
    count: result.rows.length
  };
}

/**
 * Same data for UI table (no download dialog).
 */
function getLeaveRecordsByDateRange(startDate, endDate) {
  return exportLeaveRecordsByDateRange(startDate, endDate);
}

/**
 * Build CSV string only (for download button).
 */
function buildLeaveRecordsRangeCsv(startDate, endDate) {
  var r = exportLeaveRecordsByDateRange(startDate, endDate);
  if (!r.success) return '';
  return r.csv;
}

function getLeaveRecordsInRange_(rangeStart, rangeEnd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    return { rows: [], table: [LEAVE_EXPORT_HEADERS_] };
  }

  var data = sheet.getDataRange().getValues();
  var display = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) {
    return { rows: [], table: [LEAVE_EXPORT_HEADERS_] };
  }

  var headers = data[0].map(function (h) { return String(h).trim(); });
  var idx = {
    entry: headers.indexOf('Entry Code'),
    emp: headers.indexOf('Emp ID'),
    name: headers.indexOf('Emp Name'),
    dept: headers.indexOf('Department'),
    cat: headers.indexOf('Category'),
    type: headers.indexOf('Leave Type'),
    start: headers.indexOf('Start Date'),
    end: headers.indexOf('End Date'),
    reason: headers.indexOf('Leave Reason'),
    days: headers.indexOf('No of Days'),
    util: headers.indexOf('Leave Utilized'),
    year: headers.indexOf('Entitlement Year'),
    entered: headers.indexOf('Date Entered'),
    by: headers.indexOf('Entered By')
  };

  var parseD = typeof parseLeaveDate_ === 'function' ? parseLeaveDate_ : parseExportDate_;

  var rows = [];
  var table = [LEAVE_EXPORT_HEADERS_];

  for (var i = 1; i < data.length; i++) {
    var s = idx.start >= 0
      ? (parseD(data[i][idx.start]) || parseD(display[i][idx.start]))
      : null;
    var e = idx.end >= 0
      ? (parseD(data[i][idx.end]) || parseD(display[i][idx.end]))
      : null;
    if (!s || !e) continue;

    // Overlap: leave.end >= rangeStart AND leave.start <= rangeEnd
    if (e.getTime() < rangeStart.getTime()) continue;
    if (s.getTime() > rangeEnd.getTime()) continue;

    var obj = {
      entryCode: idx.entry >= 0 ? data[i][idx.entry] : '',
      empId: idx.emp >= 0 ? String(data[i][idx.emp] || '').trim().toUpperCase() : '',
      empName: idx.name >= 0 ? data[i][idx.name] : '',
      department: idx.dept >= 0 ? data[i][idx.dept] : '',
      category: idx.cat >= 0 ? data[i][idx.cat] : '',
      leaveType: idx.type >= 0 ? data[i][idx.type] : '',
      startDate: s,
      endDate: e,
      leaveReason: idx.reason >= 0 ? data[i][idx.reason] : '',
      noOfDays: idx.days >= 0 ? data[i][idx.days] : '',
      leaveUtilized: idx.util >= 0 ? data[i][idx.util] : '',
      entitlementYear: idx.year >= 0 ? data[i][idx.year] : '',
      dateEntered: idx.entered >= 0 ? data[i][idx.entered] : '',
      enteredBy: idx.by >= 0 ? data[i][idx.by] : ''
    };
    rows.push(obj);
    table.push([
      obj.entryCode,
      obj.empId,
      obj.empName,
      obj.department,
      obj.category,
      obj.leaveType,
      formatExportDate_(obj.startDate),
      formatExportDate_(obj.endDate),
      obj.leaveReason,
      obj.noOfDays,
      obj.leaveUtilized,
      obj.entitlementYear,
      obj.dateEntered instanceof Date ? formatExportDate_(obj.dateEntered) : obj.dateEntered,
      obj.enteredBy
    ]);
  }

  // Newest start first
  rows.sort(function (a, b) { return b.startDate.getTime() - a.startDate.getTime(); });

  return { rows: rows, table: table };
}

var LEAVE_EXPORT_HEADERS_ = [
  'Entry Code', 'Emp ID', 'Emp Name', 'Department', 'Category', 'Leave Type',
  'Start Date', 'End Date', 'Leave Reason', 'No of Days', 'Leave Utilized',
  'Entitlement Year', 'Date Entered', 'Entered By'
];

function parseExportDate_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  if (typeof parseLeaveDate_ === 'function') {
    var p = parseLeaveDate_(val);
    if (p) return p;
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function formatExportDate_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2);
  }
}

/** Warm caches used by Leave Entry UI */
function preloadLeaveEntryData() {
  try {
    if (typeof loadEmployeeMapCached_ === 'function') loadEmployeeMapCached_();
    if (typeof loadPoliciesCached_ === 'function') loadPoliciesCached_();
    return { success: true, message: 'Employee list and policies loaded into memory.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
