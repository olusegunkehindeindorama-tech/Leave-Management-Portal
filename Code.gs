
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
    var dbPass = String(data[i][passCol]).trim(); // Kept case-sensitive
    var dbName = String(data[i][nameCol]).trim();
    
    if (dbUser === searchId && dbPass === String(password)) {
      var isAdmin = (dbName.toLowerCase() === 'olusegun kehinde');
      
      return { 
        success: true, 
        user: {
          id: String(data[i][idCol]).trim(), // Return original casing for records
          name: dbName,
          role: isAdmin ? 'admin' : 'standard'
        }
      };
    }
  }
  
  return { success: false, message: "Invalid User ID or Password." };
}

/**
 * Fetches employee profile and calculated balances for the Leave Entry Form
 */
function getEmployeeForForm(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblemployee');
  
  if (!empSheet) return { error: "Employee sheet not found." };
  
  var data = empSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find Employee Row
  var empRow = data.find(function(row) { 
    return String(row[0]).toLowerCase() === String(empId).toLowerCase().trim(); 
  });
  
  if (!empRow) return { error: "Employee ID not found." };
  
  // Map values (adjust index based on your actual tblemployee columns)
  // Assuming: 0: ID, 1: Name, 2: Dept, 3: Category, 4: BU
  var empData = {
    id: empRow[1],
    name: empRow[5],
    dept: empRow[6],
    category: empRow[2],
    bu: empRow[0]
  };
  
  // Use your existing backend engine to get balances
  // Note: Ensure calculateLeaveUtilized or your balance logic function is accessible
  var balances = getLeaveBalancesForEmployee(empData.id); 
  
  return {
    profile: empData,
    balances: balances // Should return list of {type, entitlement, utilized, balance}
  };
}

/**
 * Helper to get all Employee IDs for the datalist (autocomplete)
 */
function getEmployeeIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblemployee');
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
}

/**
 * Processes and appends a new leave entry into tblLeave.
 * @param {Object} formData - The data collected from the frontend.
 * @param {Object} userSession - The currently logged-in user object.
 */
function submitLeaveRequest(formData, userSession) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tblLeave = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  
  if (!tblLeave) return { success: false, message: "Error: tblLeave sheet missing." };

  var leaveData = tblLeave.getDataRange().getValues();
  var headers = leaveData[0].map(function(h) { return String(h).trim(); });
  
  // ==========================================
  // 1. GENERATE USER INITIALS
  // ==========================================
  var nameParts = String(userSession.name).trim().split(' ');
  var initials = "";
  if (nameParts.length >= 2) {
    initials = (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase();
  } else if (nameParts.length === 1 && nameParts[0] !== "") {
    initials = nameParts[0].substring(0, 2).toUpperCase();
  } else {
    initials = "XX";
  }
  
  // ==========================================
  // 2. CALCULATE NEXT SERIAL NUMBER
  // ==========================================
  var entryCodeIdx = headers.indexOf("Entry Code");
  var maxSerial = 0;
  
  for (var i = 1; i < leaveData.length; i++) {
    var code = String(leaveData[i][entryCodeIdx]).trim();
    // Regex matches uppercase letters followed by a hyphen and numbers (e.g., OK-1025)
    // This safely ignores DB entry codes if their format differs entirely
    var match = code.match(/^[A-Z]+-(\d+)$/); 
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxSerial) maxSerial = num;
    }
  }
  
  var nextSerial = maxSerial === 0 ? 1000 : maxSerial + 1;
  var newEntryCode = initials + "-" + nextSerial;
  
  // ==========================================
  // 3. FETCH LEAVE CODE FROM POLICY
  // ==========================================
  var pData = policySheet.getDataRange().getValues();
  var pHeaders = pData[0].map(function(h) { return String(h).trim(); });
  var pTypeIdx = pHeaders.indexOf("Leave Type"); 
  var pCodeIdx = pHeaders.indexOf("Leave Code"); 
  
  var sysLeaveCode = "";
  for (var p = 1; p < pData.length; p++) {
    if (String(pData[p][pTypeIdx]).trim() === formData.leaveType) {
      sysLeaveCode = pData[p][pCodeIdx];
      break;
    }
  }

  // ==========================================
  // 4. DATE AND UTILIZATION CALCULATIONS
  // ==========================================
  var sDate = new Date(formData.startDate);
  var eDate = new Date(formData.endDate);
  
  // Actual calendar days
  var noOfDays = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  var entitlementYear = sDate.getFullYear();
  
  // Hook into your granular shift calculator for Leave Utilized.
  // We wrap this in a try/catch block just in case the backend engine fails during calculation.
  var utilized = 0;
  try {
    utilized = calculateLeaveUtilize(formData.empId, sDate, eDate); 
  } catch(e) {
    utilized = noOfDays; // Fallback to raw days if script errors out
  }
  
  // Prevent submission if utilized exceeds balance 
  // (Frontend handles UI block, this is strict backend security)
  if (utilized > formData.availableBalance) {
    return { success: false, message: "Rejected: Required leave days (" + utilized + ") exceeds available balance." };
  }

  // ==========================================
  // 5. BUILD AND APPEND RECORD
  // ==========================================
  var rowObj = {
    "Entry Code": newEntryCode,
    "Leave Code": sysLeaveCode,
    "Emp ID": formData.empId,
    "Emp Name": formData.empName,
    "Dept": formData.dept,
    "Category": formData.category,
    "BU": formData.bu,
    "Leave Type": formData.leaveType,
    "Start Date": sDate,
    "End Date": eDate,
    "Leave Reason": formData.leaveReason,
    "No of Days": noOfDays,
    "Leave Utilized": utilized,
    "Entitlement Year": entitlementYear,
    "Date Entered": new Date(), // Enforces immediate system timestamp
    "Entered By": userSession.name,
    "Date Modified": "",
    "Modified By": "",
    "DB Remark": "Not Uploaded",
    "Upload Date": "",
    "Upload By": ""
  };
  
  // Map properties exactly to the column layout of tblLeave
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
  
  // Sort descending by Start Date (Most recent first)
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
  
  // Find the exact row using Entry Code
  var targetRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === updateData.entryCode) {
      targetRowIdx = i + 1; // +1 because array is 0-indexed and sheet is 1-indexed
      break;
    }
  }
  
  if (targetRowIdx === -1) return { success: false, message: "Record not found." };
  
  // Recalculate days and utilization based on new dates
  var sDate = new Date(updateData.startDate);
  var eDate = new Date(updateData.endDate);
  var noOfDays = Math.round((eDate.getTime() - sDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  
  var utilized = 0;
  try {
    utilized = calculateLeaveUtilize(updateData.empId, sDate, eDate); 
  } catch(e) {
    utilized = noOfDays; 
  }
  
  // Perform updates mapping to exact columns
  var updates = [
    { col: "Start Date", val: sDate },
    { col: "End Date", val: eDate },
    { col: "Leave Reason", val: updateData.leaveReason },
    { col: "No of Days", val: noOfDays },
    { col: "Leave Utilized", val: utilized },
    { col: "Date Modified", val: new Date() }, // Dynamic modification timestamp
    { col: "Modified By", val: userSession.name } // Logs the logged-in user
  ];
  
  updates.forEach(function(u) {
    var colIdx = headers.indexOf(u.col);
    if (colIdx > -1) {
      tblLeave.getRange(targetRowIdx, colIdx + 1).setValue(u.val);
    }
  });
  
  return { success: true, message: "Record " + updateData.entryCode + " updated successfully." };
}
