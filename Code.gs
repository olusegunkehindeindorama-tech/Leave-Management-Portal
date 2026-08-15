/** SERVE THE WEB APP */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Leave Management')
      .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function loginUser(userId, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('userData');
  if (!userSheet) return { success: false, message: "System Error: userData sheet missing." };
  var data = userSheet.getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === searchId && String(data[i][1]).trim() === String(password)) {
      var dbName = String(data[i][2]).trim();
      return { success: true, user: { id: String(data[i][0]).trim(), name: dbName, role: dbName.toLowerCase() === 'olusegun kehinde' ? 'admin' : 'standard' } };
    }
  }
  return { success: false, message: "Invalid User ID or Password." };
}

function getSecurityQuestion(userId) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('userData').getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === searchId) {
      var q = String(data[i][3] || '').trim();
      if (!q) return { success: false, message: "No security question set. Contact Mr. Olusegun Kehinde." };
      return { success: true, question: q, userId: String(data[i][0]).trim() };
    }
  }
  return { success: false, message: "User ID not found." };
}

function resetPasswordWithAnswer(userId, answer) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('userData');
  var data = sheet.getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();
  var supplied = String(answer || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === searchId) {
      if (String(data[i][4] || '').trim().toLowerCase() !== supplied)
        return { success: false, message: "Incorrect answer. Password was not changed." };
      var newPass = String(data[i][0]).trim().toLowerCase();
      sheet.getRange(i + 1, 2).setValue(newPass);
      return { success: true, message: "Password reset to lowercase User ID: " + newPass };
    }
  }
  return { success: false, message: "User ID not found." };
}

function getHeaderIndex_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/** Fast form payload: cached emp + single-emp balance */
function getEmployeeForForm(empId) {
  var search = String(empId).trim().toUpperCase();
  var empMap = loadEmployeeMapCached_();
  var emp = empMap[search];
  if (!emp) {
    // partial unique match
    var partials = [];
    Object.keys(empMap).forEach(function(id) {
      if (id === '_headers') return;
      if (id.indexOf(search) !== -1) partials.push(id);
    });
    if (partials.length === 1) emp = empMap[partials[0]];
    else if (partials.length > 1) return { error: "Multiple employees match. Select from suggestions." };
    else return { error: "Employee ID not found." };
  }

  var empData = {
    id: String(emp['Emp ID'] || search).trim().toUpperCase(),
    name: String(emp['Emp Name'] || '').trim(),
    dept: String(emp['Department'] || '').trim(),
    category: String(emp['Category'] || '').trim(),
    bu: String(emp['Business Unit'] || '').trim()
  };

  var balancePayload = apiGetEmployeeBalance(empData.id);
  if (balancePayload.error) return balancePayload;

  var balances = [];
  // Show on balance table only where Balance Page Show = Yes; dropdown gets ALL entitled
  Object.keys(balancePayload.balances || {}).forEach(function(type) {
    var d = (balancePayload.detail && balancePayload.detail[type]) ? balancePayload.detail[type] : {};
    balances.push({
      type: type,
      entitlement: balancePayload.entitlements[type],
      utilized: balancePayload.usage[type] || 0,
      balance: balancePayload.balances[type],
      detail: d,
      show: d.show !== false
    });
  });

  return {
    profile: empData,
    balances: balances,
    entitledTypes: balancePayload.entitledTypes || Object.keys(balancePayload.balances || {})
  };
}

function searchEmployees(query, limit) {
  var empMap = loadEmployeeMapCached_();
  var needle = String(query || '').trim().toLowerCase();
  var max = Number(limit) || 10;
  var matches = [];
  Object.keys(empMap).forEach(function(id) {
    if (id === '_headers' || matches.length >= max) return;
    if (id.toLowerCase().indexOf(needle) !== -1) {
      var e = empMap[id];
      matches.push({ id: id, name: String(e['Emp Name'] || ''), department: String(e['Department'] || '') });
    }
  });
  return matches;
}

function submitLeaveRequest(formData, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  if (!tblLeave) return { success: false, message: "tblLeave missing." };

  var leaveData = tblLeave.getDataRange().getValues();
  var headers = leaveData[0].map(function(h) { return String(h).trim(); });

  var nameParts = String(userSession.name).trim().split(/\s+/);
  var initials = nameParts.length >= 2
    ? (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase()
    : (nameParts[0] ? nameParts[0].substring(0, 2).toUpperCase() : 'XX');

  var entryCodeIdx = headers.indexOf('Entry Code');
  var maxSerial = 0;
  for (var i = 1; i < leaveData.length; i++) {
    var m = String(leaveData[i][entryCodeIdx]).trim().match(/^[A-Z]+-(\d+)$/);
    if (m) { var n = parseInt(m[1], 10); if (n > maxSerial) maxSerial = n; }
  }
  var newEntryCode = initials + '-' + (maxSerial === 0 ? 1000 : maxSerial + 1);

  var sysLeaveCode = '';
  var policies = loadPoliciesCached_();
  for (var p = 0; p < policies.length; p++) {
    if (String(policies[p]['Leave Type'] || '').trim() === formData.leaveType) {
      sysLeaveCode = String(policies[p]['DB Leave Code'] || '').trim();
      break;
    }
  }

  var sDate = new Date(formData.startDate);
  var eDate = new Date(formData.endDate);
  var noOfDays = Math.round((eDate - sDate) / 86400000) + 1;
  var utilized = 0;
  try { utilized = calculateLeaveUtilize(formData.empId, sDate, eDate, formData.leaveType); }
  catch (e) { utilized = noOfDays; }

  if (formData.availableBalance !== 'Unlimited' && utilized > Number(formData.availableBalance)) {
    return { success: false, message: 'Rejected: required days (' + utilized + ') exceed available balance.' };
  }

  var rowObj = {
    'Entry Code': newEntryCode, 'Leave Code': sysLeaveCode, 'Emp ID': formData.empId,
    'Emp Name': formData.empName, 'Department': formData.dept, 'Category': formData.category,
    'Leave Type': formData.leaveType, 'Start Date': sDate, 'End Date': eDate,
    'Leave Reason': formData.leaveReason, 'No of Days': noOfDays, 'Leave Utilized': utilized,
    'Entitlement Year': sDate.getFullYear(), 'Date Entered': new Date(), 'Entered By': userSession.name,
    'Date Modified': '', 'Modified By': '', 'BU': formData.bu, 'DB Remark': 'Not Uploaded',
    'Upload Date': '', 'Uploaded By': ''
  };
  tblLeave.appendRow(headers.map(function(h) { return rowObj[h] !== undefined ? rowObj[h] : ''; }));
  invalidateEmpCaches_(formData.empId);
  return { success: true, message: 'Leave recorded as ' + newEntryCode + ' (utilized ' + utilized + ' days)' };
}

function getEmployeeLeaveHistory(empId) {
  var rows = loadLeaveRowsForEmp_(empId);
  var history = rows.map(function(r) {
    return {
      entryCode: r['Entry Code'],
      empId: r['Emp ID'],
      empName: r['Emp Name'],
      type: r['Leave Type'],
      startDate: r['Start Date'],
      endDate: r['End Date'],
      noOfDays: r['No of Days'],
      entitlementYear: r['Entitlement Year'],
      utilized: r['Leave Utilized'],
      reason: r['Leave Reason'],
      status: r['DB Remark']
    };
  });
  history.sort(function(a, b) { return new Date(b.startDate) - new Date(a.startDate); });
  return history;
}

function updateLeaveRecord(updateData, userSession) {
  var tblLeave = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  var data = tblLeave.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var entryCodeIdx = headers.indexOf('Entry Code');
  var target = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === updateData.entryCode) { target = i + 1; break; }
  }
  if (target < 0) return { success: false, message: 'Record not found.' };

  var sDate = new Date(updateData.startDate);
  var eDate = new Date(updateData.endDate);
  var noOfDays = Math.round((eDate - sDate) / 86400000) + 1;
  var type = String(data[target - 1][headers.indexOf('Leave Type')] || '');
  var utilized = 0;
  try { utilized = calculateLeaveUtilize(updateData.empId, sDate, eDate, type); }
  catch (e) { utilized = noOfDays; }

  [{ col: 'Start Date', val: sDate }, { col: 'End Date', val: eDate }, { col: 'Leave Reason', val: updateData.leaveReason },
   { col: 'No of Days', val: noOfDays }, { col: 'Leave Utilized', val: utilized },
   { col: 'Date Modified', val: new Date() }, { col: 'Modified By', val: userSession.name }
  ].forEach(function(u) {
    var c = headers.indexOf(u.col);
    if (c > -1) tblLeave.getRange(target, c + 1).setValue(u.val);
  });
  invalidateEmpCaches_(updateData.empId);
  return { success: true, message: 'Record ' + updateData.entryCode + ' updated.' };
}

function buildBalanceReportCsv(buFilter, deptFilter) {
  return arrayToCsv_(generateReportArray(buFilter || null, deptFilter || null));
}

function getLeaveRecordsFiltered(filters) {
  filters = filters || {};
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave').getDataRange().getValues();
  var headers = data.shift().map(function(h) { return String(h).trim(); });
  var idx = {
    entry: headers.indexOf('Entry Code'), emp: headers.indexOf('Emp ID'), name: headers.indexOf('Emp Name'),
    dept: headers.indexOf('Department'), bu: headers.indexOf('BU'), type: headers.indexOf('Leave Type'),
    start: headers.indexOf('Start Date'), end: headers.indexOf('End Date'), util: headers.indexOf('Leave Utilized'),
    remark: headers.indexOf('DB Remark'), enteredBy: headers.indexOf('Entered By')
  };
  var fromD = filters.fromDate ? new Date(filters.fromDate) : null;
  var toD = filters.toDate ? new Date(filters.toDate) : null;
  var limit = Number(filters.limit) || 500;
  var rows = [];
  for (var i = 0; i < data.length && rows.length < limit; i++) {
    var r = data[i];
    var empId = String(r[idx.emp] || '').trim().toUpperCase();
    if (filters.empId && empId.indexOf(String(filters.empId).toUpperCase()) === -1) continue;
    if (filters.bu && String(r[idx.bu]).toUpperCase() !== String(filters.bu).toUpperCase()) continue;
    if (filters.dept && String(r[idx.dept]).toUpperCase() !== String(filters.dept).toUpperCase()) continue;
    if (filters.leaveType && String(r[idx.type]).trim() !== String(filters.leaveType).trim()) continue;
    if (filters.dbRemark && String(r[idx.remark]).toLowerCase().indexOf(String(filters.dbRemark).toLowerCase()) === -1) continue;
    var sDate = new Date(r[idx.start]);
    if (fromD && !isNaN(fromD) && (isNaN(sDate) || sDate < fromD)) continue;
    if (toD && !isNaN(toD) && (isNaN(sDate) || sDate > toD)) continue;
    rows.push({ entryCode: r[idx.entry], empId: empId, empName: r[idx.name], department: r[idx.dept], bu: r[idx.bu],
      leaveType: r[idx.type], startDate: r[idx.start], endDate: r[idx.end], utilized: r[idx.util],
      dbRemark: r[idx.remark], enteredBy: r[idx.enteredBy] });
  }
  return { rows: rows, total: rows.length };
}

function buildLeaveRecordsCsv(filters) {
  var result = getLeaveRecordsFiltered(filters || {});
  var out = [['Entry Code', 'Emp ID', 'Emp Name', 'Department', 'BU', 'Leave Type', 'Start Date', 'End Date', 'Leave Utilized', 'DB Remark', 'Entered By']];
  result.rows.forEach(function(r) {
    out.push([r.entryCode, r.empId, r.empName, r.department, r.bu, r.leaveType,
      r.startDate instanceof Date ? Utilities.formatDate(r.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : r.startDate,
      r.endDate instanceof Date ? Utilities.formatDate(r.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : r.endDate,
      r.utilized, r.dbRemark, r.enteredBy]);
  });
  return arrayToCsv_(out);
}

/** Shift calendar + leave overlays for the month */
function getShiftCalendar(empId, year, month) {
  year = Number(year) || new Date().getFullYear();
  month = Number(month) || (new Date().getMonth() + 1);
  var target = String(empId).trim().toUpperCase();
  var shiftMap = loadShiftMapForEmp_(target);
  var leaveRows = loadLeaveRowsForEmp_(target);

  var leaveByDate = {};
  leaveRows.forEach(function(r) {
    var s = new Date(r['Start Date']);
    var e = new Date(r['End Date']);
    if (isNaN(s) || isNaN(e)) return;
    var cur = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    var end = new Date(e.getFullYear(), e.getMonth(), e.getDate());
    while (cur <= end) {
      if (cur.getFullYear() === year && (cur.getMonth() + 1) === month) {
        var k = formatDateKey(cur);
        if (!leaveByDate[k]) leaveByDate[k] = [];
        leaveByDate[k].push(String(r['Leave Type'] || 'Leave'));
      }
      cur.setDate(cur.getDate() + 1);
    }
  });

  var daysInMonth = new Date(year, month, 0).getDate();
  var days = [];
  for (var day = 1; day <= daysInMonth; day++) {
    var dt = new Date(year, month - 1, day);
    var key = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    days.push({
      date: key,
      day: day,
      dayOfWeek: dt.getDay(),
      shift: shiftMap[key] || '',
      leaves: leaveByDate[key] || []
    });
  }
  return { year: year, month: month, empId: target, days: days };
}

function updateShiftCode(empId, dateStr, shiftCode, userSession) {
  var shiftSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblShift');
  if (!shiftSheet) return { success: false, message: 'tblShift not found.' };
  var data = shiftSheet.getDataRange().getValues();
  var target = String(empId).trim().toUpperCase();
  var targetKey = Utilities.formatDate(new Date(dateStr), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() !== target) continue;
    var d = new Date(data[i][1]);
    if (!isNaN(d) && Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') === targetKey) { found = i + 1; break; }
  }
  var code = String(shiftCode || '').trim().toUpperCase();
  if (found > -1) {
    if (!code) { shiftSheet.deleteRow(found); invalidateEmpCaches_(target); return { success: true, message: 'Shift cleared' }; }
    shiftSheet.getRange(found, 3).setValue(code);
    invalidateEmpCaches_(target);
    return { success: true, message: 'Shift updated to ' + code };
  }
  if (!code) return { success: true, message: 'Nothing to clear' };
  shiftSheet.appendRow([target, targetKey, code]);
  invalidateEmpCaches_(target);
  return { success: true, message: 'Shift added: ' + code };
}

function arrayToCsv_(rows) {
  return rows.map(function(row) {
    return row.map(function(cell) {
      var cellStr = String(cell == null ? '' : cell);
      if (/[",\n]/.test(cellStr)) return '"' + cellStr.replace(/"/g, '""') + '"';
      return cellStr;
    }).join(',');
  }).join('\n');
}
