/**
 * UI API helpers: recent leaves, delete leave, balance CSV, multi-filter.
 */

/** Newest leave records across all employees (when no emp selected). */
function getRecentLeaveRecords(limit) {
  limit = Number(limit) || 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblLeave');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var idx = {
    entry: headers.indexOf('Entry Code'),
    emp: headers.indexOf('Emp ID'),
    name: headers.indexOf('Emp Name'),
    type: headers.indexOf('Leave Type'),
    start: headers.indexOf('Start Date'),
    end: headers.indexOf('End Date'),
    days: headers.indexOf('No of Days'),
    year: headers.indexOf('Entitlement Year'),
    reason: headers.indexOf('Leave Reason')
  };
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      entryCode: idx.entry >= 0 ? data[i][idx.entry] : '',
      empId: idx.emp >= 0 ? String(data[i][idx.emp] || '').trim().toUpperCase() : '',
      empName: idx.name >= 0 ? data[i][idx.name] : '',
      type: idx.type >= 0 ? data[i][idx.type] : '',
      startDate: typeof formatDateOnly_ === 'function' ? formatDateOnly_(data[i][idx.start]) : String(data[i][idx.start] || ''),
      endDate: typeof formatDateOnly_ === 'function' ? formatDateOnly_(data[i][idx.end]) : String(data[i][idx.end] || ''),
      noOfDays: idx.days >= 0 ? data[i][idx.days] : '',
      entitlementYear: idx.year >= 0 ? data[i][idx.year] : '',
      reason: idx.reason >= 0 ? data[i][idx.reason] : ''
    });
  }
  rows.sort(function (a, b) {
    return String(b.startDate || '').localeCompare(String(a.startDate || ''));
  });
  return rows.slice(0, limit);
}

/** Delete a leave row by Entry Code. */
function deleteLeaveRecord(entryCode, userSession) {
  entryCode = String(entryCode || '').trim();
  if (!entryCode) return { success: false, message: 'Missing entry code.' };
  var tblLeave = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  if (!tblLeave) return { success: false, message: 'tblLeave missing.' };
  var data = tblLeave.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var entryCodeIdx = headers.indexOf('Entry Code');
  var empIdx = headers.indexOf('Emp ID');
  var target = -1;
  var empId = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === entryCode) {
      target = i + 1;
      empId = empIdx >= 0 ? String(data[i][empIdx] || '').trim().toUpperCase() : '';
      break;
    }
  }
  if (target < 0) return { success: false, message: 'Record not found.' };
  tblLeave.deleteRow(target);
  if (empId && typeof invalidateEmpCaches_ === 'function') invalidateEmpCaches_(empId);
  return { success: true, message: 'Deleted ' + entryCode + '.' };
}

/**
 * Balance CSV for filtered BU/Dept (pipe-separated multi values).
 * Reads tblEmployee leave columns — does NOT recompute (fast).
 */
function generateReportArray(buFilter, deptFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet || empSheet.getLastRow() < 2) {
    return [['Emp ID', 'Emp Name', 'Business Unit', 'Department', 'Category', 'Message'],
            ['', '', '', '', '', 'No employees']];
  }
  var data = empSheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var idIdx = headers.indexOf('Emp ID');
  var nameIdx = headers.indexOf('Emp Name');
  var buIdx = headers.indexOf('Business Unit');
  var deptIdx = headers.indexOf('Department');
  var catIdx = headers.indexOf('Category');
  var genderIdx = headers.indexOf('Gender');
  var dojIdx = headers.indexOf('Date of Join');

  // Leave balance columns commonly present
  var leaveCols = ['Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'];
  var leaveIdx = {};
  leaveCols.forEach(function (c) { leaveIdx[c] = headers.indexOf(c); });

  function multiMatch(cell, filterStr) {
    if (!filterStr) return true;
    var cellU = String(cell || '').trim().toUpperCase();
    var parts = String(filterStr).split('|').map(function (s) {
      return s.trim().toUpperCase();
    }).filter(Boolean);
    if (!parts.length) return true;
    // "ALL" means no filter
    if (parts.length === 1 && parts[0] === 'ALL') return true;
    return parts.indexOf(cellU) !== -1;
  }

  var outHeaders = ['Emp ID', 'Emp Name', 'Business Unit', 'Department', 'Category', 'Gender', 'Date of Join'];
  leaveCols.forEach(function (c) {
    if (leaveIdx[c] >= 0) outHeaders.push(c);
  });
  var out = [outHeaders];

  for (var i = 1; i < data.length; i++) {
    var bu = buIdx >= 0 ? data[i][buIdx] : '';
    var dept = deptIdx >= 0 ? data[i][deptIdx] : '';
    if (!multiMatch(bu, buFilter)) continue;
    if (!multiMatch(dept, deptFilter)) continue;
    var row = [
      idIdx >= 0 ? data[i][idIdx] : '',
      nameIdx >= 0 ? data[i][nameIdx] : '',
      bu,
      dept,
      catIdx >= 0 ? data[i][catIdx] : '',
      genderIdx >= 0 ? data[i][genderIdx] : '',
      dojIdx >= 0 ? (data[i][dojIdx] instanceof Date
        ? Utilities.formatDate(data[i][dojIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : data[i][dojIdx]) : ''
    ];
    leaveCols.forEach(function (c) {
      if (leaveIdx[c] >= 0) row.push(data[i][leaveIdx[c]]);
    });
    out.push(row);
  }
  return out;
}

/** Override buildBalanceReportCsv to always return non-empty CSV string. */
function buildBalanceReportCsv(buFilter, deptFilter) {
  var arr = generateReportArray(buFilter || null, deptFilter || null);
  return arrayToCsv_(arr);
}

/** Export employee history (or recent) as CSV string for UI download. */
function buildEmployeeLeaveHistoryCsv(empId) {
  var rows;
  if (empId && String(empId).trim()) {
    rows = getEmployeeLeaveHistory(String(empId).trim().toUpperCase());
  } else {
    rows = getRecentLeaveRecords(200);
  }
  var out = [['Entry Code', 'Emp ID', 'Emp Name', 'Start Date', 'End Date', 'No of Days', 'Entitlement Year', 'Leave Type', 'Leave Reason']];
  (rows || []).forEach(function (r) {
    out.push([
      r.entryCode, r.empId, r.empName, r.startDate, r.endDate,
      r.noOfDays, r.entitlementYear, r.type, r.reason || ''
    ]);
  });
  return arrayToCsv_(out);
}
