
/**
 * IMPORT DARWINBOX LEAVES (Dynamic Column & Strict Policy Mapping)
 * Optimized for automated background triggers.
 */
function importDarwinBoxLeaves() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folderId = "1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1";

  var leaveSheet = ss.getSheetByName('tblLeave');
  var empSheet = ss.getSheetByName('tblEmployee');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');

  if (!leaveSheet || !empSheet || !policySheet) {
    console.error("Error: Missing required sheets ('tblLeave', 'tblEmployee', or 'Sys_LeavePolicies').");
    return;
  }

  // ==========================================
  // 1. CSV HEADER CONFIGURATION
  // ==========================================
  var csvHeadConfig = {
    empId: "Employee Id",
    empName: "Employee Name",
    startDate: "Leave From Date",
    endDate: "Leave To Date",
    status: "Status",
    appliedOn: "Applied On",
    leaveType: "Leave Type",
    comment: "Employee Comment"
  };

  // 2. Get the DarwinBox CSV file
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName("Leave_Application.csv");
  if (!files.hasNext()) {
    console.error("Error: File 'Leave_Application.csv' not found in the specified folder.");
    return;
  }
  var file = files.next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

  // Extract and clean CSV headers (removes trailing spaces)
  var rawCsvHeaders = csvData.shift();
  var csvHeaders = rawCsvHeaders.map(function(h) { return String(h).trim(); });

  // Find dynamic CSV Column Indices
  var cEmpIdIdx = csvHeaders.indexOf(csvHeadConfig.empId);
  var cNameIdx = csvHeaders.indexOf(csvHeadConfig.empName);
  var cStartIdx = csvHeaders.indexOf(csvHeadConfig.startDate);
  var cEndIdx = csvHeaders.indexOf(csvHeadConfig.endDate);
  var cStatusIdx = csvHeaders.indexOf(csvHeadConfig.status);
  var cAppliedIdx = csvHeaders.indexOf(csvHeadConfig.appliedOn);
  var cTypeIdx = csvHeaders.indexOf(csvHeadConfig.leaveType);
  var cCommentIdx = csvHeaders.indexOf(csvHeadConfig.comment);

  if (cEmpIdIdx === -1 || cStartIdx === -1 || cTypeIdx === -1) {
    console.error("Error: Could not find required columns in the CSV. Please check the 'csvHeadConfig' names.");
    return;
  }

  // 3. Load Policy Map (Cleaned & Trimmed)
  var policyData = policySheet.getDataRange().getValues();
  var rawPHeaders = policyData.shift();
  var pHeaders = rawPHeaders.map(function(h) { return String(h).trim(); });

  var pDBTypeIdx = pHeaders.indexOf("DB Leave Name");
  var pDBCodeIdx = pHeaders.indexOf("DB Leave Code");
  var pStdTypeIdx = pHeaders.indexOf("Leave Type");

  if (pDBTypeIdx === -1 || pStdTypeIdx === -1) {
    console.error("Error: Could not find 'DB Leave Name' or 'Leave Type' in Sys_LeavePolicies headers.");
    return;
  }

  var policyMap = {};
  for (var p = 0; p < policyData.length; p++) {
    var dbName = String(policyData[p][pDBTypeIdx]).trim().toLowerCase();
    if (dbName) {
      policyMap[dbName] = {
        stdType: String(policyData[p][pStdTypeIdx]).trim(), // Keeps full written terms as preferred
        dbCode: pDBCodeIdx !== -1 ? String(policyData[p][pDBCodeIdx]).trim() : ""
      };
    }
  }

  // 4. Load Employee Map (Cleaned)
  var empData = empSheet.getDataRange().getValues();
  var eHeaders = empData.shift().map(function(h) { return String(h).trim(); });
  var eIdIdx = eHeaders.indexOf("Emp ID");
  var eBuIdx = eHeaders.indexOf("Business Unit");
  var eCatIdx = eHeaders.indexOf("Category");
  var eDeptIdx = eHeaders.indexOf("Department");

  var empMap = {};
  for (var e = 0; e < empData.length; e++) {
    var eId = String(empData[e][eIdIdx]).trim().toUpperCase();
    if (eId) {
      empMap[eId] = {
        bu: eBuIdx !== -1 ? String(empData[e][eBuIdx]) : "",
        cat: eCatIdx !== -1 ? String(empData[e][eCatIdx]) : "",
        dept: eDeptIdx !== -1 ? String(empData[e][eDeptIdx]) : ""
      };
    }
  }

  // 5. Load tblLeave into Memory & Build Deduplication Dictionary
  var rawLeaveData = leaveSheet.getDataRange().getValues();
  var lHeaders = rawLeaveData[0].map(function(h) { return String(h).trim(); });
  var lEmpIdx = lHeaders.indexOf("Emp ID");
  var lStartIdx = lHeaders.indexOf("Start Date");
  var lEndIdx = lHeaders.indexOf("End Date");

  var leaveDataMemory = [];
  var existingKeys = {};

  for (var r = 0; r < rawLeaveData.length; r++) {
    var row = rawLeaveData[r];
    if (r > 0 && String(row[lEmpIdx]).trim() === "") continue;

    leaveDataMemory.push(row);

    if (r > 0) {
      var empIdStr = String(row[lEmpIdx]).trim().toUpperCase();
      var sDate = new Date(row[lStartIdx]);
      var eDate = new Date(row[lEndIdx]);

      if (empIdStr && !isNaN(sDate.getTime()) && !isNaN(eDate.getTime())) {
        // Relies on formatDateKey being available globally in the project
        var key = empIdStr + "_" + formatDateKey(sDate) + "_" + formatDateKey(eDate);
        existingKeys[key] = true;
      }
    }
  }

  // 6. Process CSV dynamically and Push New Rows
  var newRowsCount = 0;

  for (var i = 0; i < csvData.length; i++) {
    var row = csvData[i];

    if (cStatusIdx !== -1 && String(row[cStatusIdx]).trim() !== "Approved") continue;

    var empId = String(row[cEmpIdIdx]).trim().toUpperCase();
    var startDate = new Date(row[cStartIdx]);
    var endDate = new Date(row[cEndIdx]);

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) continue;

    var fingerprintKey = empId + "_" + formatDateKey(startDate) + "_" + formatDateKey(endDate);

    if (!existingKeys[fingerprintKey]) {
      var empInfo = empMap[empId] || {bu: "", cat: "", dept: ""};

      // Perform strict Policy Lookup
      var dbLeaveTypeRaw = String(row[cTypeIdx]).trim();
      var mappedPolicy = policyMap[dbLeaveTypeRaw.toLowerCase()];

      // If no match found in Sys_LeavePolicies, fallback to original CSV strings to avoid blanks
      if (!mappedPolicy) {
        mappedPolicy = { stdType: dbLeaveTypeRaw, dbCode: "" };
      }

      var newRow = [];
      newRow[lHeaders.indexOf("Entry Code")] = "DB-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
      newRow[lHeaders.indexOf("Leave Code")] = mappedPolicy.dbCode;
      newRow[lHeaders.indexOf("Emp ID")] = empId;
      newRow[lHeaders.indexOf("Emp Name")] = cNameIdx !== -1 ? row[cNameIdx] : "";
      newRow[lHeaders.indexOf("Department")] = empInfo.dept;
      newRow[lHeaders.indexOf("Category")] = empInfo.cat;
      newRow[lHeaders.indexOf("Leave Type")] = mappedPolicy.stdType;
      newRow[lHeaders.indexOf("Start Date")] = startDate;
      newRow[lHeaders.indexOf("End Date")] = endDate;
      newRow[lHeaders.indexOf("Leave Reason")] = cCommentIdx !== -1 ? row[cCommentIdx] : "";
      newRow[lHeaders.indexOf("No of Days")] = "";
      newRow[lHeaders.indexOf("Leave Utilized")] = "";
      newRow[lHeaders.indexOf("Entitlement Year")] = "";
      newRow[lHeaders.indexOf("Date Entered")] = cAppliedIdx !== -1 ? new Date(row[cAppliedIdx]) : "";
      newRow[lHeaders.indexOf("Entered By")] = "Darwinbox";
      newRow[lHeaders.indexOf("Date Modified")] = "";
      newRow[lHeaders.indexOf("Modified By")] = "";
      newRow[lHeaders.indexOf("BU")] = empInfo.bu;
      newRow[lHeaders.indexOf("DB Remark")] = "Original";
      newRow[lHeaders.indexOf("Upload Date")] = new Date();
      var uploadByIdx = lHeaders.indexOf("Uploaded By") !== -1 ? lHeaders.indexOf("Uploaded By") : lHeaders.indexOf("Upload By");
      if (uploadByIdx !== -1) newRow[uploadByIdx] = "Automation";

      for (var col = 0; col < lHeaders.length; col++) {
        if (newRow[col] === undefined) newRow[col] = "";
      }

      leaveDataMemory.push(newRow);
      existingKeys[fingerprintKey] = true;
      newRowsCount++;
    }
  }

  // 7. Bulk Write with Capacity Checking
  if (newRowsCount > 0) {
    var requiredRows = leaveDataMemory.length;
    var maxRows = leaveSheet.getMaxRows();

    if (requiredRows > maxRows) {
      leaveSheet.insertRowsAfter(maxRows, requiredRows - maxRows + 10);
    }

    leaveSheet.clearContents();
    leaveSheet.getRange(1, 1, leaveDataMemory.length, leaveDataMemory[0].length).setValues(leaveDataMemory);

    console.log(newRowsCount + " new DB records imported successfully.");
    return { success: true, message: newRowsCount + " new DB records imported successfully." };
  } else {
    console.log("No new approved leave records found to import.");
    return { success: true, message: "No new approved leave records found to import." };
  }
}
