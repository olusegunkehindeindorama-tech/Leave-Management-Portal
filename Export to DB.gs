
/**
 * Backend endpoint for the HTML UI. It returns CSV text to the browser so the
 * frontend can trigger a local download without a server-side modal dialog.
 */
function buildDarwinBoxExportCsv() {
  var result = buildDarwinBoxExportData_();
  if (!result.success) return result;
  return {
    success: true,
    message: result.records + " records compiled successfully.",
    records: result.records,
    fileName: "DarwinBox_Upload.csv",
    csv: arrayToCsv_(result.rows)
  };
}

/**
 * Compatibility wrapper for manual Apps Script testing.
 */
function exportToDarwinBox() {
  var result = buildDarwinBoxExportCsv();
  if (!result.success) {
    SpreadsheetApp.getUi().alert(result.message);
    return;
  }

  var encodedCsv = Utilities.base64Encode(Utilities.newBlob(result.csv).getBytes());
  var html =
    '<html><body>' +
    '<h3 style="font-family: sans-serif; color: #333;">Export Ready</h3>' +
    '<p style="font-family: sans-serif; font-size: 14px;">' + result.records + ' records compiled successfully.</p>' +
    '<a href="data:text/csv;base64,' + encodedCsv + '" download="' + result.fileName + '" ' +
    'style="display: inline-block; padding: 10px 20px; background-color: #1a73e8; color: white; text-decoration: none; border-radius: 4px; font-family: sans-serif; font-weight: bold;" ' +
    'onclick="setTimeout(function(){ google.script.host.close(); }, 500);">Download CSV File</a>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(320).setHeight(160), "DarwinBox Export");
}

/**
 * Creates the DarwinBox upload rows from tblLeave according to README rules.
 */
function buildDarwinBoxExportData_() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveSheet = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var folderId = "1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1";

  if (!leaveSheet || !policySheet) {
    return { success: false, message: "Error: Missing 'tblLeave' or 'Sys_LeavePolicies' sheet." };
  }

  var existingDBKeys = {};
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName("Leave_Application.csv");

  if (files.hasNext()) {
    var csvData = Utilities.parseCsv(files.next().getBlob().getDataAsString());
    if (csvData.length > 0) {
      var csvHeaders = csvData.shift().map(function(h) { return String(h).trim(); });
      var cEmpIdIdx = csvHeaders.indexOf("Employee Id");
      var cStartIdx = csvHeaders.indexOf("Leave From Date");
      var cEndIdx = csvHeaders.indexOf("Leave To Date");

      for (var i = 0; i < csvData.length; i++) {
        var row = csvData[i];
        var eId = cEmpIdIdx === -1 ? "" : String(row[cEmpIdIdx]).trim().toUpperCase();
        var sDate = cStartIdx === -1 ? null : new Date(row[cStartIdx]);
        var eDate = cEndIdx === -1 ? null : new Date(row[cEndIdx]);
        if (eId && sDate && eDate && !isNaN(sDate.getTime()) && !isNaN(eDate.getTime())) {
          existingDBKeys[eId + "_" + formatDateKey(sDate) + "_" + formatDateKey(eDate)] = true;
        }
      }
    }
  } else {
    console.warn("Leave_Application.csv not found. All matched records will be exported.");
  }

  var policyData = policySheet.getDataRange().getValues();
  var pHeaders = policyData.shift().map(function(h) { return String(h).trim(); });
  var pDBNameIdx = pHeaders.indexOf("DB Leave Name");
  var pDBCodeIdx = pHeaders.indexOf("DB Leave Code");
  var codeToNameMap = {};

  for (var p = 0; p < policyData.length; p++) {
    var dbCode = pDBCodeIdx === -1 ? "" : String(policyData[p][pDBCodeIdx]).trim();
    var dbName = pDBNameIdx === -1 ? "" : String(policyData[p][pDBNameIdx]).trim();
    if (dbCode) codeToNameMap[dbCode] = dbName;
  }

  var leaveData = leaveSheet.getDataRange().getValues();
  var lHeaders = leaveData.shift().map(function(h) { return String(h).trim(); });
  var lEmpIdx = lHeaders.indexOf("Emp ID");
  var lCodeIdx = lHeaders.indexOf("Leave Code");
  var lReasonIdx = lHeaders.indexOf("Leave Reason");
  var lStartIdx = lHeaders.indexOf("Start Date");
  var lEndIdx = lHeaders.indexOf("End Date");
  var lDateEnteredIdx = lHeaders.indexOf("Date Entered");

  var twoMonthsAgo = new Date();
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);

  var exportRows = [[
    "Email/Employee ID", "Leave name", "Leave code", "Subcategory",
    "Ispaid/unpaid", "Leave Message", "From Date", "To Date",
    "Is half Day?", "Revoke leave", "Leave reason"
  ]];

  for (var l = 0; l < leaveData.length; l++) {
    var leaveRow = leaveData[l];
    var empId = lEmpIdx === -1 ? "" : String(leaveRow[lEmpIdx]).trim().toUpperCase();
    var leaveCode = lCodeIdx === -1 ? "" : String(leaveRow[lCodeIdx]).trim();
    var startDate = lStartIdx === -1 ? null : new Date(leaveRow[lStartIdx]);
    var endDate = lEndIdx === -1 ? null : new Date(leaveRow[lEndIdx]);
    var dateEntered = lDateEnteredIdx === -1 ? null : new Date(leaveRow[lDateEnteredIdx]);

    if (!empId || !startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) continue;
    if (!dateEntered || isNaN(dateEntered.getTime()) || dateEntered < twoMonthsAgo) continue;
    if (!/^(FRT|SR|IFF)/.test(empId)) continue;
    if (leaveCode === "" || leaveCode.toUpperCase() === "NA") continue;

    var fingerprintKey = empId + "_" + formatDateKey(startDate) + "_" + formatDateKey(endDate);
    if (existingDBKeys[fingerprintKey]) continue;

    exportRows.push([
      empId,
      codeToNameMap[leaveCode] || "",
      leaveCode,
      "",
      "Paid",
      lReasonIdx === -1 ? "" : String(leaveRow[lReasonIdx]).trim(),
      formatDateForDBExport(startDate),
      formatDateForDBExport(endDate),
      "No",
      "No",
      "Personal"
    ]);
  }

  if (exportRows.length <= 1) {
    return { success: false, message: "No new records found for export in the last 2 months." };
  }

  return { success: true, records: exportRows.length - 1, rows: exportRows };
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
