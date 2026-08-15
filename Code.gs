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
 * Checks credentials against 'userData' sheet.
 * User ID is case-insensitive, Password is case-sensitive.
 * Admin = Full Name "Olusegun Kehinde"
 */
function loginUser(userId, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('userData');

  if (!userSheet) return { success: false, message: "System Error: userData sheet missing." };

  var data = userSheet.getDataRange().getValues();
  var idCol = 0;
  var passCol = 1;
  var nameCol = 2;

  var searchId = String(userId).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var dbUser = String(data[i][idCol]).trim().toLowerCase();
    var dbPass = String(data[i][passCol]).trim();
    var dbName = String(data[i][nameCol]).trim();

    if (dbUser === searchId && dbPass === String(password)) {
      var isAdmin = (dbName.toLowerCase() === 'olusegun kehinde');

      return {
        success: true,
        user: {
          id: String(data[i][idCol]).trim(),
          name: dbName,
          role: isAdmin ? 'admin' : 'standard'
        }
      };
    }
  }

  return { success: false, message: "Invalid User ID or Password." };
}

/**
 * Finds sheet header indexes by any accepted header name.
 */
function getHeaderIndex_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Safely returns a row value by header aliases.
 */
function getRowValue_(row, headers, names) {
  var idx = getHeaderIndex_(headers, names);
  return idx === -1 ? "" : row[idx];
}

/**
 * Fetches employee profile and calculated balances for the Leave Entry Form.
 * Accepts either an exact ID or a partial ID that uniquely matches an employee.
 */
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
    if (rowId === search) {
      exactMatch = data[i];
      break;
    }
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
  var balanceTypes = Object.keys(balancePayload.balances || {});
  for (var b = 0; b < balanceTypes.length; b++) {
    var type = balanceTypes[b];
    balances.push({
      type: type,
      entitlement: balancePayload.entitlements[type],
      utilized: balancePayload.usage[type] || 0,
      balance: balancePayload.balances[type]
    });
  }

  return {
    profile: empData,
    balances: balances
  };
}

/**
 * Partial employee search for the UI autocomplete.
 */
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

/**
 * Helper to get all Employee IDs for the datalist (autocomplete)
 * Uses live sheet name tblEmployee and Emp ID column.
 */
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

/**
 * Processes and appends a new leave entry into tblLeave.
 * Maps exactly to live sheet headers:
 * Entry Code, Leave Code, Emp ID, Emp Name, Department, Category, Leave Type,
 * Start Date, End Date, Leave Reason, No of Days, Leave Utilized, Entitlement Year,
 * Date Entered, Entered By, Date Modified, Modified By, BU, DB Remark, Upload Date, Uploaded By
 *
 * Leave Code is looked up from Sys_LeavePolicies."DB Leave Code" by Leave Type.
 */
function submitLeaveRequest(formData, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');

  if (!tblLeave) return { success: false, message: "Error: tblLeave sheet missing." };

  var leaveData = tblLeave.getDataRange().getValues();
  var headers = leaveData[0].map(function(h) { return String(h).trim(); });

  // 1. GENERATE USER INITIALS
  var nameParts = String(userSession.name).trim().split(/\s+/);
  var initials = "";
  if (nameParts.length >= 2) {
    initials = (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase();
  } else if (nameParts.length === 1 && nameParts[0] !== "") {
    initials = nameParts[0].substring(0, 2).toUpperCase();
  } else {
    initials = "XX";
  }

  // 2. CALCULATE NEXT SERIAL NUMBER (ignore DB- prefixed codes)
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

  var nextSerial = maxSerial === 0 ? 1000 : maxSerial + 1;
  var newEntryCode = initials + "-" + nextSerial;

  // 3. FETCH LEAVE CODE FROM POLICY (live header is "DB Leave Code")
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

  // 4. DATE AND UTILIZATION CALCULATIONS
  var sDate = new Date(formData.startDate);
  var eDate = new Date(formData.endDate);

  var noOfDays = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  var entitlementYear = sDate.getFullYear();

  var utilized = 0;
  try {
    utilized = calculateLeaveUtilize(formData.empId, sDate, eDate);
  } catch (e) {
    utilized = noOfDays;
  }

  if (utilized > formData.availableBalance) {
    return { success: false, message: "Rejected: Required leave days (" + utilized + ") exceeds available balance." };
  }

  // 5. BUILD AND APPEND RECORD — keys must match live tblLeave headers exactly
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
    var field = rowObj[h];
    return field !== undefined ? field : "";
  });

  tblLeave.appendRow(rowToAppend);

  return { success: true, message: "Leave successfully recorded as " + newEntryCode };
}

/**
 * Fetches leave history for a specific employee, sorted descending by Start Date.
 */
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

  history.sort(function(a, b) {
    return new Date(b.startDate) - new Date(a.startDate);
  });

  return history;
}

/**
 * Updates an existing leave record and tags the modifier.
 */
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
  try {
    utilized = calculateLeaveUtilize(updateData.empId, sDate, eDate);
  } catch (e) {
    utilized = noOfDays;
  }

  var updates = [
    { col: "Start Date", val: sDate },
    { col: "End Date", val: eDate },
    { col: "Leave Reason", val: updateData.leaveReason },
    { col: "No of Days", val: noOfDays },
    { col: "Leave Utilized", val: utilized },
    { col: "Date Modified", val: new Date() },
    { col: "Modified By", val: userSession.name }
  ];

  updates.forEach(function(u) {
    var colIdx = headers.indexOf(u.col);
    if (colIdx > -1) {
      tblLeave.getRange(targetRowIdx, colIdx + 1).setValue(u.val);
    }
  });

  return { success: true, message: "Record " + updateData.entryCode + " updated successfully." };
}

/**
 * Builds a downloadable balance report CSV for the Reports module.
 */
function buildBalanceReportCsv() {
  var report = generateReportArray();
  return arrayToCsv_(report);
}

/**
 * Shared CSV serializer used by UI download endpoints.
 */
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
