/**
 * Refresh Employee Leave balance + year Entitlement columns.
 * Use from UI button (Reports + Admin).
 */
function refreshEmployeeLeaveBalances() {
  var result = updateAllEmployeeLeaveBalances();
  if (!result || !result.success) return result;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
    var leaveSheet = ss.getSheetByName('tblLeave');
    var policySheet = ss.getSheetByName('Sys_LeavePolicies');
    var sbSheet = ss.getSheetByName('StartingBal');
    if (!empSheet || !policySheet) return result;

    var today = new Date();
    var currentYear = today.getFullYear();
    var prevYear = currentYear - 1;

    var policies = [];
    var pData = policySheet.getDataRange().getValues();
    if (pData.length > 1) {
      var pH = pData[0].map(function (h) { return String(h).trim(); });
      for (var i = 1; i < pData.length; i++) {
        var o = {};
        for (var c = 0; c < pH.length; c++) o[pH[c]] = pData[i][c];
        policies.push(o);
      }
    }

    var startingBalMap = {};
    if (sbSheet) {
      var sbData = sbSheet.getDataRange().getValues();
      if (sbData.length > 1) {
        var sbH = sbData[0].map(function (h) { return String(h).trim(); });
        var eIdx = sbH.indexOf('Emp No');
        if (eIdx < 0) eIdx = sbH.indexOf('Emp ID');
        var bIdx = sbH.indexOf(prevYear + ' Balance');
        if (eIdx >= 0 && bIdx >= 0) {
          for (var s = 1; s < sbData.length; s++) {
            var sid = String(sbData[s][eIdx] || '').trim().toUpperCase();
            if (sid) startingBalMap[sid] = Number(sbData[s][bIdx]) || 0;
          }
        }
      }
    }

    var leaveUsageByEmp = {};
    if (leaveSheet) {
      var lData = leaveSheet.getDataRange().getValues();
      if (lData.length > 1) {
        var lH = lData[0].map(function (h) { return String(h).trim(); });
        var liEmp = lH.indexOf('Emp ID');
        var liType = lH.indexOf('Leave Type');
        var liUtil = lH.indexOf('Leave Utilized');
        var liYear = lH.indexOf('Entitlement Year');
        var liStart = lH.indexOf('Start Date');
        for (var r = 1; r < lData.length; r++) {
          var eid = liEmp >= 0 ? String(lData[r][liEmp] || '').trim().toUpperCase() : '';
          var lt = liType >= 0 ? String(lData[r][liType] || '').trim() : '';
          var util = liUtil >= 0 ? Number(lData[r][liUtil]) || 0 : 0;
          var ey = liYear >= 0 ? Number(lData[r][liYear]) : 0;
          if (!ey && liStart >= 0) {
            var sd = lData[r][liStart];
            if (sd instanceof Date && !isNaN(sd.getTime())) ey = sd.getFullYear();
          }
          if (!eid || !lt) continue;
          if (!leaveUsageByEmp[eid]) leaveUsageByEmp[eid] = {};
          if (!leaveUsageByEmp[eid][lt]) leaveUsageByEmp[eid][lt] = {};
          if (!leaveUsageByEmp[eid][lt][ey]) leaveUsageByEmp[eid][lt][ey] = 0;
          leaveUsageByEmp[eid][lt][ey] += util;
        }
      }
    }

    var eData = empSheet.getDataRange().getValues();
    var eH = eData[0].map(function (h) { return String(h).trim(); });
    var employees = [];
    for (var er = 1; er < eData.length; er++) {
      var rowObj = {};
      for (var ec = 0; ec < eH.length; ec++) rowObj[eH[ec]] = eData[er][ec];
      rowObj._sheetRow = er + 1;
      employees.push(rowObj);
    }

    if (typeof writeEntitlementYearColumns_ === 'function') {
      writeEntitlementYearColumns_(empSheet, employees, policies, leaveUsageByEmp, startingBalMap, today);
      SpreadsheetApp.flush();
    }

    result.message = (result.message || 'Done') + ' Entitlement year columns updated.';
  } catch (e) {
    result.message = (result.message || '') + ' (Entitlement cols: ' + e.message + ')';
  }
  return result;
}
