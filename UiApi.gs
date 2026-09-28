/**
 * UI API helpers: recent leaves, soft-delete, balance CSV, entitlement export.
 */

function getRecentLeaveRecords(limit) {
  limit = Number(limit) || 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblLeave');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var display = sh.getDataRange().getDisplayValues();
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
      startDate: safeDateOnlyFromSheet_(data[i][idx.start], display[i][idx.start]),
      endDate: safeDateOnlyFromSheet_(data[i][idx.end], display[i][idx.end]),
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

function safeDateOnlyFromSheet_(val, displayVal) {
  if (typeof formatDateOnly_ === 'function') {
    var f = formatDateOnly_(val);
    if (f) return f;
  }
  var d = String(displayVal || '').trim();
  if (typeof parseDateOnly_ === 'function' && d) {
    var p = parseDateOnly_(d);
    if (p && typeof formatDateOnly_ === 'function') return formatDateOnly_(p);
  }
  return d.substring(0, 10);
}

function deleteLeaveRecord(entryCode, userSession) {
  entryCode = String(entryCode || '').trim();
  if (!entryCode) return { success: false, message: 'Missing entry code.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  if (!tblLeave) return { success: false, message: 'tblLeave missing.' };

  var data = tblLeave.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var entryCodeIdx = headers.indexOf('Entry Code');
  var empIdx = headers.indexOf('Emp ID');
  var target = -1;
  var empId = '';
  var rowValues = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === entryCode) {
      target = i + 1;
      empId = empIdx >= 0 ? String(data[i][empIdx] || '').trim().toUpperCase() : '';
      rowValues = data[i].slice();
      break;
    }
  }
  if (target < 0) return { success: false, message: 'Record not found.' };

  var arch = ss.getSheetByName('tblLeaveDeleted');
  if (!arch) {
    arch = ss.insertSheet('tblLeaveDeleted');
    arch.appendRow(headers.concat(['Deleted At', 'Deleted By']));
  } else if (arch.getLastRow() === 0) {
    arch.appendRow(headers.concat(['Deleted At', 'Deleted By']));
  } else {
    var aH = arch.getRange(1, 1, 1, arch.getLastColumn()).getValues()[0].map(function (h) {
      return String(h).trim();
    });
    if (aH.indexOf('Deleted At') < 0) {
      arch.getRange(1, aH.length + 1).setValue('Deleted At');
      arch.getRange(1, aH.length + 2).setValue('Deleted By');
    }
  }

  var deletedBy = (userSession && userSession.name) ? userSession.name : '';
  var deletedAt = new Date();
  var archiveRow = rowValues.concat([deletedAt, deletedBy]);
  var archCols = arch.getLastColumn();
  while (archiveRow.length < archCols) archiveRow.push('');
  if (archiveRow.length > archCols) archiveRow = archiveRow.slice(0, archCols);
  arch.appendRow(archiveRow);

  tblLeave.deleteRow(target);
  if (empId && typeof invalidateEmpCaches_ === 'function') invalidateEmpCaches_(empId);
  return { success: true, message: 'Moved ' + entryCode + ' to tblLeaveDeleted.' };
}

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

  var leaveCols = ['Annual', 'Casual', 'Compassionate', 'Examination', 'Study', 'Maternity', 'Probation'];
  var leaveIdx = {};
  leaveCols.forEach(function (c) { leaveIdx[c] = headers.indexOf(c); });

  var today = new Date();
  var cy = today.getFullYear();
  var py = cy - 1;
  var entCols = [
    py + ' Entitlement', py + ' Utilized', py + ' Balance',
    cy + ' Entitlement', cy + ' Utilized', cy + ' Balance'
  ];
  var entIdx = {};
  entCols.forEach(function (c) { entIdx[c] = headers.indexOf(c); });

  function multiMatch(cell, filterStr) {
    if (!filterStr) return true;
    var cellU = String(cell || '').trim().toUpperCase();
    var parts = String(filterStr).split('|').map(function (s) {
      return s.trim().toUpperCase();
    }).filter(Boolean);
    if (!parts.length || (parts.length === 1 && parts[0] === 'ALL')) return true;
    return parts.indexOf(cellU) !== -1;
  }

  var outHeaders = ['Emp ID', 'Emp Name', 'Business Unit', 'Department', 'Category', 'Gender', 'Date of Join'];
  leaveCols.forEach(function (c) {
    if (leaveIdx[c] >= 0) outHeaders.push(c);
  });
  entCols.forEach(function (c) {
    if (entIdx[c] >= 0) outHeaders.push(c);
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
      bu, dept,
      catIdx >= 0 ? data[i][catIdx] : '',
      genderIdx >= 0 ? data[i][genderIdx] : '',
      dojIdx >= 0 ? (typeof formatDateOnly_ === 'function'
        ? formatDateOnly_(data[i][dojIdx])
        : data[i][dojIdx]) : ''
    ];
    leaveCols.forEach(function (c) {
      if (leaveIdx[c] >= 0) row.push(data[i][leaveIdx[c]]);
    });
    entCols.forEach(function (c) {
      if (entIdx[c] >= 0) row.push(data[i][entIdx[c]]);
    });
    out.push(row);
  }
  return out;
}

function buildBalanceReportCsv(buFilter, deptFilter) {
  var arr = generateReportArray(buFilter || null, deptFilter || null);
  return arrayToCsv_(arr);
}

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

function exportEntitlementBalanceCsv(buFilter, deptFilter) {
  var arr = generateEntitlementBalanceArray_(buFilter || null, deptFilter || null);
  return arrayToCsv_(arr);
}

function generateEntitlementBalanceArray_(buFilter, deptFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet || empSheet.getLastRow() < 2) {
    return [['Message'], ['No employees']];
  }
  var data = empSheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var today = new Date();
  var cy = today.getFullYear();
  var py = cy - 1;

  var detailCols = ['Emp ID', 'Emp Name', 'Business Unit', 'Department', 'Category', 'Gender', 'Date of Join'];
  var pyCols = [py + ' Entitlement', py + ' Utilized', py + ' Balance'];
  var cyCols = [cy + ' Entitlement', cy + ' Utilized', cy + ' Balance'];
  var includePrev = false;
  pyCols.forEach(function (c) {
    if (headers.indexOf(c) >= 0) includePrev = true;
  });

  var outH = detailCols.slice();
  if (includePrev) outH = outH.concat(pyCols);
  outH = outH.concat(cyCols);
  var out = [outH];

  function multiMatch(cell, filterStr) {
    if (!filterStr) return true;
    var cellU = String(cell || '').trim().toUpperCase();
    var parts = String(filterStr).split('|').map(function (s) {
      return s.trim().toUpperCase();
    }).filter(Boolean);
    if (!parts.length || (parts.length === 1 && parts[0] === 'ALL')) return true;
    return parts.indexOf(cellU) !== -1;
  }

  var hIdx = {};
  headers.forEach(function (h, i) { hIdx[h] = i; });

  for (var i = 1; i < data.length; i++) {
    var bu = hIdx['Business Unit'] != null ? data[i][hIdx['Business Unit']] : '';
    var dept = hIdx['Department'] != null ? data[i][hIdx['Department']] : '';
    if (!multiMatch(bu, buFilter)) continue;
    if (!multiMatch(dept, deptFilter)) continue;
    var row = [];
    detailCols.forEach(function (c) {
      var v = hIdx[c] != null ? data[i][hIdx[c]] : '';
      if (c === 'Date of Join' && typeof formatDateOnly_ === 'function') v = formatDateOnly_(v) || v;
      row.push(v);
    });
    if (includePrev) {
      pyCols.forEach(function (c) {
        row.push(hIdx[c] != null ? data[i][hIdx[c]] : '');
      });
    }
    cyCols.forEach(function (c) {
      row.push(hIdx[c] != null ? data[i][hIdx[c]] : '');
    });
    out.push(row);
  }
  return out;
}
