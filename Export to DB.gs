
/**
 * EXPORT LOCAL LEAVES TO DARWINBOX CSV
 * Generates a mapped and filtered CSV of local tblLeave records 
 * entered within the last 2 months that do not exist in the DarwinBox CSV, 
 * specifically for FRT, SR, and IFF employees.
 */
function exportToDarwinBox() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var leaveSheet = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var folderId = "1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1"; // Same folder used for import
  
  if (!leaveSheet || !policySheet) {
    SpreadsheetApp.getUi().alert("Error: Missing 'tblLeave' or 'Sys_LeavePolicies' sheet.");
    return;
  }
  
  // ==========================================
  // 1. BUILD FINGERPRINTS FROM EXISTING DARWINBOX FILE
  // ==========================================
  var existingDBKeys = {};
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName("Leave_Application.csv");
  
  if (files.hasNext()) {
    var file = files.next();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
    var csvHeaders = csvData.shift().map(function(h) { return String(h).trim(); });
    
    var cEmpIdIdx = csvHeaders.indexOf("Employee Id");
    var cStartIdx = csvHeaders.indexOf("Leave From Date");
    var cEndIdx = csvHeaders.indexOf("Leave To Date");
    
    for (var i = 0; i < csvData.length; i++) {
      var row = csvData[i];
      var eId = String(row[cEmpIdIdx]).trim().toUpperCase();
      var sDate = new Date(row[cStartIdx]);
      var eDate = new Date(row[cEndIdx]);
      
      if (eId && !isNaN(sDate.getTime()) && !isNaN(eDate.getTime())) {
        var key = eId + "_" + formatDateKey(sDate) + "_" + formatDateKey(eDate);
        existingDBKeys[key] = true;
      }
    }
  } else {
    console.warn("Leave_Application.csv not found. All matched records will be exported.");
  }
  
  // ==========================================
  // 2. LOAD POLICY MAP (Leave Code -> DB Leave Name)
  // ==========================================
  var policyData = policySheet.getDataRange().getValues();
  var pHeaders = policyData.shift().map(function(h) { return String(h).trim(); });
  
  var pDBNameIdx = pHeaders.indexOf("DB Leave Name");
  var pDBCodeIdx = pHeaders.indexOf("DB Leave Code");
  
  var codeToNameMap = {};
  for (var p = 0; p < policyData.length; p++) {
    var dbCode = String(policyData[p][pDBCodeIdx]).trim();
    var dbName = String(policyData[p][pDBNameIdx]).trim();
    if (dbCode) {
      codeToNameMap[dbCode] = dbName;
    }
  }

  // ==========================================
  // 3. PROCESS TBLLEAVE & BUILD EXPORT ARRAY
  // ==========================================
  var leaveData = leaveSheet.getDataRange().getValues();
  var lHeaders = leaveData.shift().map(function(h) { return String(h).trim(); });
  
  var lEmpIdx = lHeaders.indexOf("Emp ID");
  var lCodeIdx = lHeaders.indexOf("Leave Code"); 
  var lReasonIdx = lHeaders.indexOf("Leave Reason");
  var lStartIdx = lHeaders.indexOf("Start Date");
  var lEndIdx = lHeaders.indexOf("End Date");
  var lDateEnteredIdx = lHeaders.indexOf("Date Entered"); // Get Date Entered column
  
  // Calculate the cutoff date (Exactly 2 months ago dynamically)
  var twoMonthsAgo = new Date();
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
  
  // DarwinBox target headers
  var exportData = [[
    "Email/Employee ID", "Leave name", "Leave code", "Subcategory", 
    "Ispaid/unpaid", "Leave Message", "From Date", "To Date", 
    "Is half Day?", "Revoke leave", "Leave reason"
  ]];
  
  for (var l = 0; l < leaveData.length; l++) {
    var row = leaveData[l];
    var empId = String(row[lEmpIdx]).trim().toUpperCase();
    var leaveCode = String(row[lCodeIdx]).trim();
    var startDate = new Date(row[lStartIdx]);
    var endDate = new Date(row[lEndIdx]);
    var dateEntered = new Date(row[lDateEnteredIdx]);
    
    // Skip empty rows or invalid dates
    if (!empId || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) continue;
    
    // NEW Filter: Must be entered within the last 2 months
    if (isNaN(dateEntered.getTime()) || dateEntered < twoMonthsAgo) continue;
    
    // Filter 1: Must start with FRT, SR, or IFF
    var isTargetEmp = empId.indexOf("FRT") === 0 || empId.indexOf("SR") === 0 || empId.indexOf("IFF") === 0;
    if (!isTargetEmp) continue;
    
    // Filter 2: Leave Code must not be NA or blank
    if (leaveCode.toUpperCase() === "NA" || leaveCode === "") continue;
    
    // Filter 3: Must not exist in DarwinBox already
    var fingerprintKey = empId + "_" + formatDateKey(startDate) + "_" + formatDateKey(endDate);
    if (existingDBKeys[fingerprintKey]) continue;
    
    // Perform Lookup for Leave Name
    var dbLeaveName = codeToNameMap[leaveCode] || "";
    
    // Build the mapped row
    exportData.push([
      empId,                                      
      dbLeaveName,                                
      leaveCode,                                  
      "",                                         
      "Paid",                                     
      String(row[lReasonIdx]).trim(),             
      formatDateForDBExport(startDate),           
      formatDateForDBExport(endDate),             
      "No",                                       
      "No",                                       
      "Personal"                                  
    ]);
  }
  
  // ==========================================
  // 4. GENERATE CSV & TRIGGER DOWNLOAD UI
  // ==========================================
  if (exportData.length <= 1) {
    SpreadsheetApp.getUi().alert("No new records found for export in the last 2 months.");
    return;
  }
  
  var csvString = exportData.map(function(row) {
    return row.map(function(cell) {
      var cellStr = String(cell);
      if (cellStr.indexOf(',') > -1 || cellStr.indexOf('\n') > -1 || cellStr.indexOf('"') > -1) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',');
  }).join('\n');
  
  var encodedCsv = Utilities.base64Encode(Utilities.newBlob(csvString).getBytes());
  var html = 
    '<html><body>' +
    '<h3 style="font-family: sans-serif; color: #333;">Export Ready</h3>' +
    '<p style="font-family: sans-serif; font-size: 14px;">' + (exportData.length - 1) + ' records compiled successfully.</p>' +
    '<a href="data:text/csv;base64,' + encodedCsv + '" download="DarwinBox_Upload.csv" ' +
    'style="display: inline-block; padding: 10px 20px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-family: sans-serif; font-weight: bold;" ' +
    'onclick="setTimeout(function(){ google.script.host.close(); }, 500);">' +
    'Download CSV File</a>' +
    '</body></html>';
    
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "DarwinBox Export");
}

/** 
 * HELPER: Formats Date for the DarwinBox CSV Export (DD/MM/YYYY)
 */
function formatDateForDBExport(dateObj) {
  var d = ("0" + dateObj.getDate()).slice(-2);
  var m = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  var y = dateObj.getFullYear();
  return d + "/" + m + "/" + y; 
}