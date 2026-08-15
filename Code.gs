/**
 * SERVE THE WEB APP
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Leave Management')
      .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * AUTHENTICATE USER
 * User ID case-insensitive, Password case-sensitive.
 * Admin = Full Name "Olusegun Kehinde"
 */
function loginUser(userId, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('userData');
  if (!userSheet) return { success: false, message: "System Error: userData sheet missing." };

  var data = userSheet.getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var dbUser = String(data[i][0]).trim().toLowerCase();
    var dbPass = String(data[i][1]).trim();
    var dbName = String(data[i][2]).trim();

    if (dbUser === searchId && dbPass === String(password)) {
      return {
        success: true,
        user: {
          id: String(data[i][0]).trim(),
          name: dbName,
          role: (dbName.toLowerCase() === 'olusegun kehinde') ? 'admin' : 'standard'
        }
      };
    }
  }
  return { success: false, message: "Invalid User ID or Password." };
}

/**
 * Returns the security question for a user (for self-service password reset).
 * userData columns: User ID | Password | Full Name | Question | Answer
 */
function getSecurityQuestion(userId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('userData');
  if (!userSheet) return { success: false, message: "userData sheet missing." };

  var data = userSheet.getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === searchId) {
      var question = String(data[i][3] || "").trim();
      if (!question) {
        return { success: false, message: "No security question set for this account. Contact Mr. Olusegun Kehinde." };
      }
      return { success: true, question: question, userId: String(data[i][0]).trim() };
    }
  }
  return { success: false, message: "User ID not found." };
}

/**
 * Resets password to lowercase(userId) if the answer matches (case-insensitive).
 */
function resetPasswordWithAnswer(userId, answer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('userData');
  if (!userSheet) return { success: false, message: "userData sheet missing." };

  var data = userSheet.getDataRange().getValues();
  var searchId = String(userId).trim().toLowerCase();
  var supplied = String(answer || "").trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === searchId) {
      var storedAnswer = String(data[i][4] || "").trim().toLowerCase();
      if (!storedAnswer) {
        return { success: false, message: "No security answer on file. Contact Mr. Olusegun Kehinde." };
      }
      if (storedAnswer !== supplied) {
        return { success: false, message: "Incorrect answer. Password was not changed." };
      }
      var originalId = String(data[i][0]).trim();
      var newPass = originalId.toLowerCase();
      userSheet.getRange(i + 1, 2).setValue(newPass); // Password column B
      return {
        success: true,
        message: "Password reset successfully. Your new password is the lowercase of your User ID (" + newPass + "). Please log in and change it if needed."
      };
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

function getRowValue_(row, headers, names) {
  var idx = getHeaderIndex_(headers, names);
  return idx === -1 ? "" : row[idx];
}

function getEmployeeForForm(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) return { error: "Employee sheet not found." };

  var data = empSheet.getDataRange().getValues();
  if (data.length < 2) return { error: "Employee sheet has no employee records." };

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idIdx = getHeaderIndex_(headers, ["Emp ID", "Employee ID", "Emp No"]);
  if (idIdx === -1) return { error: "Employee ID column not found on tblEmployee." };

  var search = String(empId).trim().toUpperCase();
  var exactMatch = null;
  var partialMatches = [];

  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][idIdx]).trim().toUpperCase();
    if (!rowId) continue;
    if (rowId === search) { exactMatch = data[i]; break; }
    if (rowId.toLowerCase().indexOf(search.toLowerCase()) !== -1) partialMatches.push(data[i]);
  }

  var empRow = exactMatch || (partialMatches.length === 1 ? partialMatches[0] : null);
  if (!empRow) {
    return { error: partialMatches.length > 1 ? "Multiple employees match that text. Please select one from the suggestions." : "Employee ID not found." };
  }

  var empData = {
    id: String(getRowValue_(empRow, headers, ["Emp ID", "Employee ID", "Emp No"])).trim().toUpperCase(),
    name: String(getRowValue_(empRow, headers, ["Emp Name", "Employee Name", "Name"])).trim(),
    dept: String(getRowValue_(empRow, headers, ["Department", "Dept"])).trim(),
    category: String(getRowValue_(empRow, headers, ["Category"])).trim(),
    bu: String(getRowValue_(empRow, headers, ["Business Unit", "BU"])).trim()
  };

  var balancePayload = apiGetEmployeeBalance(empData.id);
  if (balancePayload.error) return balancePayload;

  var balances = [];
  Object.keys(balancePayload.balances || {}).forEach(function(type) {
    balances.push({
      type: type,
      entitlement: balancePayload.entitlements[type],
      utilized: balancePayload.usage[type] || 0,
      balance: balancePayload.balances[type],
      detail: (balancePayload.detail && balancePayload.detail[type]) ? balancePayload.detail[type] : null
    });
  });

  return { profile: empData, balances: balances };
}

function searchEmployees(query, limit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) return [];

  var data = empSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idIdx = getHeaderIndex_(headers, ["Emp ID", "Employee ID", "Emp No"]);
  var nameIdx = getHeaderIndex_(headers, ["Emp Name", "Employee Name", "Name"]);
  var deptIdx = getHeaderIndex_(headers, ["Department", "Dept"]);
  if (idIdx === -1) return [];

  var needle = String(query || "").trim().toLowerCase();
  var max = Number(limit) || 10;
  var matches = [];

  for (var i = 1; i < data.length && matches.length < max; i++) {
    var id = String(data[i][idIdx]).trim();
    if (!id) continue;
    if (id.toLowerCase().indexOf(needle) !== -1) {
      matches.push({
        id: id,
        name: nameIdx === -1 ? "" : String(data[i][nameIdx]).trim(),
        department: deptIdx === -1 ? "" : String(data[i][deptIdx]).trim()
      });
    }
  }
  return matches;
}

function getEmployeeIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idIdx = getHeaderIndex_(headers, ["Emp ID", "Employee ID", "Emp No"]);
  if (idIdx === -1) return [];
  var ids = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idIdx]).trim();
    if (id) ids.push(id);
  }
  return ids;
}

function submitLeaveRequest(formData, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  if (!tblLeave) return { success: false, message: "Error: tblLeave sheet missing." };

  var leaveData = tblLeave.getDataRange().getValues();
  var headers = leaveData[0].map(function(h) { return String(h).trim(); });

  var nameParts = String(userSession.name).trim().split(/\s+/);
  var initials = nameParts.length >= 2
    ? (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase()
    : (nameParts[0] ? nameParts[0].substring(0, 2).toUpperCase() : "XX");

  var entryCodeIdx = headers.indexOf("Entry Code");
  var maxSerial = 0;
  for (var i = 1; i < leaveData.length; i++) {
    var code = String(leaveData[i][entryCodeIdx]).trim();
    var match = code.match(/^[A-Z]+-(\d+)$/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxSerial) maxSerial = num;
    }
  }
  var newEntryCode = initials + "-" + (maxSerial === 0 ? 1000 : maxSerial + 1);

  var sysLeaveCode = "";
  if (policySheet) {
    var pData = policySheet.getDataRange().getValues();
    var pHeaders = pData[0].map(function(h) { return String(h).trim(); });
    var pTypeIdx = pHeaders.indexOf("Leave Type");
    var pCodeIdx = pHeaders.indexOf("DB Leave Code");
    if (pCodeIdx === -1) pCodeIdx = pHeaders.indexOf("Leave Code");
    for (var p = 1; p < pData.length; p++) {
      if (String(pData[p][pTypeIdx]).trim() === formData.leaveType) {
        sysLeaveCode = pCodeIdx !== -1 ? String(pData[p][pCodeIdx]).trim() : "";
        break;
      }
    }
  }

  var sDate = new Date(formData.startDate);
  var eDate = new Date(formData.endDate);
  var noOfDays = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  var entitlementYear = sDate.getFullYear();

  var utilized = 0;
  try { utilized = calculateLeaveUtilize(formData.empId, sDate, eDate); }
  catch (e) { utilized = noOfDays; }

  if (utilized > formData.availableBalance) {
    return { success: false, message: "Rejected: Required leave days (" + utilized + ") exceeds available balance." };
  }

  var rowObj = {
    "Entry Code": newEntryCode,
    "Leave Code": sysLeaveCode,
    "Emp ID": formData.empId,
    "Emp Name": formData.empName,
    "Department": formData.dept,
    "Category": formData.category,
    "Leave Type": formData.leaveType,
    "Start Date": sDate,
    "End Date": eDate,
    "Leave Reason": formData.leaveReason,
    "No of Days": noOfDays,
    "Leave Utilized": utilized,
    "Entitlement Year": entitlementYear,
    "Date Entered": new Date(),
    "Entered By": userSession.name,
    "Date Modified": "",
    "Modified By": "",
    "BU": formData.bu,
    "DB Remark": "Not Uploaded",
    "Upload Date": "",
    "Uploaded By": ""
  };

  var rowToAppend = headers.map(function(h) {
    return rowObj[h] !== undefined ? rowObj[h] : "";
  });
  tblLeave.appendRow(rowToAppend);
  return { success: true, message: "Leave successfully recorded as " + newEntryCode };
}

function getEmployeeLeaveHistory(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  if (!tblLeave) return [];

  var data = tblLeave.getDataRange().getValues();
  var headers = data.shift().map(function(h) { return String(h).trim(); });

  var empIdx = headers.indexOf("Emp ID");
  var entryCodeIdx = headers.indexOf("Entry Code");
  var typeIdx = headers.indexOf("Leave Type");
  var startIdx = headers.indexOf("Start Date");
  var endIdx = headers.indexOf("End Date");
  var reasonIdx = headers.indexOf("Leave Reason");
  var statusIdx = headers.indexOf("DB Remark");

  var history = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][empIdx]).trim().toUpperCase() === String(empId).trim().toUpperCase()) {
      history.push({
        entryCode: data[i][entryCodeIdx],
        type: data[i][typeIdx],
        startDate: data[i][startIdx],
        endDate: data[i][endIdx],
        reason: data[i][reasonIdx],
        status: data[i][statusIdx]
      });
    }
  }
  history.sort(function(a, b) { return new Date(b.startDate) - new Date(a.startDate); });
  return history;
}

function updateLeaveRecord(updateData, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  var data = tblLeave.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var entryCodeIdx = headers.indexOf("Entry Code");

  var targetRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === updateData.entryCode) {
      targetRowIdx = i + 1;
      break;
    }
  }
  if (targetRowIdx === -1) return { success: false, message: "Record not found." };

  var sDate = new Date(updateData.startDate);
  var eDate = new Date(updateData.endDate);
  var noOfDays = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  var utilized = 0;
  try { utilized = calculateLeaveUtilize(updateData.empId, sDate, eDate); }
  catch (e) { utilized = noOfDays; }

  [
    { col: "Start Date", val: sDate },
    { col: "End Date", val: eDate },
    { col: "Leave Reason", val: updateData.leaveReason },
    { col: "No of Days", val: noOfDays },
    { col: "Leave Utilized", val: utilized },
    { col: "Date Modified", val: new Date() },
    { col: "Modified By", val: userSession.name }
  ].forEach(function(u) {
    var colIdx = headers.indexOf(u.col);
    if (colIdx > -1) tblLeave.getRange(targetRowIdx, colIdx + 1).setValue(u.val);
  });

  return { success: true, message: "Record " + updateData.entryCode + " updated successfully." };
}

function buildBalanceReportCsv(buFilter, deptFilter) {
  return arrayToCsv_(generateReportArray(buFilter || null, deptFilter || null));
}

/**
 * Filterable leave records for Reports module.
 * filters: { empId, bu, dept, leaveType, fromDate, toDate, dbRemark, limit }
 */
function getLeaveRecordsFiltered(filters) {
  filters = filters || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  if (!tblLeave) return { rows: [], total: 0 };

  var data = tblLeave.getDataRange().getValues();
  var headers = data.shift().map(function(h) { return String(h).trim(); });

  var idx = {
    entry: headers.indexOf("Entry Code"),
    emp: headers.indexOf("Emp ID"),
    name: headers.indexOf("Emp Name"),
    dept: headers.indexOf("Department"),
    bu: headers.indexOf("BU"),
    type: headers.indexOf("Leave Type"),
    start: headers.indexOf("Start Date"),
    end: headers.indexOf("End Date"),
    util: headers.indexOf("Leave Utilized"),
    remark: headers.indexOf("DB Remark"),
    enteredBy: headers.indexOf("Entered By")
  };

  var fromD = filters.fromDate ? new Date(filters.fromDate) : null;
  var toD = filters.toDate ? new Date(filters.toDate) : null;
  var limit = Number(filters.limit) || 500;
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var empId = String(r[idx.emp] || "").trim().toUpperCase();
    if (filters.empId && empId.indexOf(String(filters.empId).trim().toUpperCase()) === -1) continue;
    if (filters.bu && String(r[idx.bu] || "").trim().toUpperCase() !== String(filters.bu).trim().toUpperCase()) continue;
    if (filters.dept && String(r[idx.dept] || "").trim().toUpperCase() !== String(filters.dept).trim().toUpperCase()) continue;
    if (filters.leaveType && String(r[idx.type] || "").trim() !== String(filters.leaveType).trim()) continue;
    if (filters.dbRemark && String(r[idx.remark] || "").trim().toLowerCase().indexOf(String(filters.dbRemark).trim().toLowerCase()) === -1) continue;

    var sDate = new Date(r[idx.start]);
    if (fromD && !isNaN(fromD.getTime()) && (isNaN(sDate.getTime()) || sDate < fromD)) continue;
    if (toD && !isNaN(toD.getTime()) && (isNaN(sDate.getTime()) || sDate > toD)) continue;

    rows.push({
      entryCode: r[idx.entry],
      empId: empId,
      empName: r[idx.name],
      department: r[idx.dept],
      bu: r[idx.bu],
      leaveType: r[idx.type],
      startDate: r[idx.start],
      endDate: r[idx.end],
      utilized: r[idx.util],
      dbRemark: r[idx.remark],
      enteredBy: r[idx.enteredBy]
    });

    if (rows.length >= limit) break;
  }

  return { rows: rows, total: rows.length, headers: headers };
}

function buildLeaveRecordsCsv(filters) {
  var result = getLeaveRecordsFiltered(filters || {});
  var out = [["Entry Code", "Emp ID", "Emp Name", "Department", "BU", "Leave Type", "Start Date", "End Date", "Leave Utilized", "DB Remark", "Entered By"]];
  result.rows.forEach(function(r) {
    out.push([
      r.entryCode, r.empId, r.empName, r.department, r.bu, r.leaveType,
      r.startDate instanceof Date ? Utilities.formatDate(r.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : r.startDate,
      r.endDate instanceof Date ? Utilities.formatDate(r.endDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : r.endDate,
      r.utilized, r.dbRemark, r.enteredBy
    ]);
  });
  return arrayToCsv_(out);
}

/**
 * Shift calendar for one employee + month.
 * Returns { year, month, days: [{ date, shift, dayOfWeek }] }
 */
function getShiftCalendar(empId, year, month) {
  // month is 1-12
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName('tblShift');
  if (!shiftSheet) return { error: "tblShift not found", days: [] };

  year = Number(year) || new Date().getFullYear();
  month = Number(month) || (new Date().getMonth() + 1);

  var data = shiftSheet.getDataRange().getValues();
  if (data.length < 2) return { year: year, month: month, empId: empId, days: [] };

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var empIdx = getHeaderIndex_(headers, ["Emp ID", "Employee ID", "Emp No"]);
  var dateIdx = getHeaderIndex_(headers, ["Date", "Shift Date"]);
  var codeIdx = getHeaderIndex_(headers, ["Shift", "Shift Code", "Code"]);
  if (empIdx === -1) empIdx = 0;
  if (dateIdx === -1) dateIdx = 1;
  if (codeIdx === -1) codeIdx = 2;

  var target = String(empId).trim().toUpperCase();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][empIdx]).trim().toUpperCase() !== target) continue;
    var d = new Date(data[i][dateIdx]);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== year || (d.getMonth() + 1) !== month) continue;
    map[Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")] = String(data[i][codeIdx]).trim();
  }

  var daysInMonth = new Date(year, month, 0).getDate();
  var days = [];
  for (var day = 1; day <= daysInMonth; day++) {
    var dt = new Date(year, month - 1, day);
    var key = Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd");
    days.push({
      date: key,
      day: day,
      dayOfWeek: dt.getDay(), // 0=Sun
      shift: map[key] || ""
    });
  }

  return { year: year, month: month, empId: target, days: days };
}

/**
 * Update or insert a single shift cell for an employee on a date.
 */
function updateShiftCode(empId, dateStr, shiftCode, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName('tblShift');
  if (!shiftSheet) return { success: false, message: "tblShift not found." };

  var data = shiftSheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var empIdx = getHeaderIndex_(headers, ["Emp ID", "Employee ID", "Emp No"]);
  var dateIdx = getHeaderIndex_(headers, ["Date", "Shift Date"]);
  var codeIdx = getHeaderIndex_(headers, ["Shift", "Shift Code", "Code"]);
  if (empIdx === -1) empIdx = 0;
  if (dateIdx === -1) dateIdx = 1;
  if (codeIdx === -1) codeIdx = 2;

  var target = String(empId).trim().toUpperCase();
  var targetDate = new Date(dateStr);
  if (isNaN(targetDate.getTime())) return { success: false, message: "Invalid date." };

  var targetKey = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][empIdx]).trim().toUpperCase() !== target) continue;
    var d = new Date(data[i][dateIdx]);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") === targetKey) {
      foundRow = i + 1;
      break;
    }
  }

  var code = String(shiftCode || "").trim().toUpperCase();

  if (foundRow > -1) {
    if (code === "") {
      shiftSheet.deleteRow(foundRow);
      return { success: true, message: "Shift cleared for " + targetKey };
    }
    shiftSheet.getRange(foundRow, codeIdx + 1).setValue(code);
    return { success: true, message: "Shift updated to " + code + " on " + targetKey };
  }

  if (code === "") return { success: true, message: "No existing shift to clear." };

  // Append new row — keep column order Emp ID, Date, Shift
  var newRow = ["", "", ""];
  newRow[empIdx] = target;
  newRow[dateIdx] = targetKey;
  newRow[codeIdx] = code;
  // Ensure length 3
  while (newRow.length < 3) newRow.push("");
  shiftSheet.appendRow(newRow.slice(0, 3));
  return { success: true, message: "Shift added: " + code + " on " + targetKey };
}

function arrayToCsv_(rows) {
  return rows.map(function(row) {
    return row.map(function(cell) {
      var cellStr = String(cell === null || cell === undefined ? "" : cell);
      if (cellStr.indexOf(',') > -1 || cellStr.indexOf('\n') > -1 || cellStr.indexOf('"') > -1) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',');
  }).join('\n');
}
