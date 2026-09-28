/**
 * Precomputed Entitlement / Utilized / Balance columns on tblEmployee.
 * This-year Entitlement = Annual + Probation entitlements
 * This-year Utilized = Annual + Casual + Examination + Probation (current year)
 * Balance = Entitlement - Utilized
 * Same for previous year when carry-forward is active.
 */

function ensureEntitlementYearColumns_(empSheet, headers, yearList) {
  var needed = [];
  yearList.forEach(function (y) {
    needed.push(y + ' Entitlement');
    needed.push(y + ' Utilized');
    needed.push(y + ' Balance');
  });
  var lastCol = empSheet.getLastColumn();
  needed.forEach(function (name) {
    if (headers.indexOf(name) < 0) {
      lastCol++;
      empSheet.getRange(1, lastCol).setValue(name);
      headers.push(name);
    }
  });
  return headers;
}

function writeEntitlementYearColumns_(empSheet, employees, policies, leaveUsageByEmp, startingBalMap, today) {
  today = today || new Date();
  var cy = today.getFullYear();
  var py = cy - 1;

  var headers = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  headers = ensureEntitlementYearColumns_(empSheet, headers, [py, cy]);

  var colIdx = {};
  headers.forEach(function (h, i) { colIdx[h] = i; });

  var carryActive = isCarryForwardActive_(policies, today);

  employees.forEach(function (emp) {
    var empId = String(emp['Emp ID'] || '').trim().toUpperCase();
    var usage = leaveUsageByEmp[empId] || {};
    var pack = {
      empId: empId,
      bu: String(emp['Business Unit'] || '').trim(),
      category: String(emp['Category'] || '').trim(),
      status: String(emp['Status'] || '').trim(),
      gender: String(emp['Gender'] || '').trim(),
      doj: emp['Date of Join'],
      carryGross: Number(startingBalMap[empId]) || 0
    };

    var core = (typeof computeBalancesForEmployeeCore_ === 'function')
      ? computeBalancesForEmployeeCore_(pack, policies, usage, today)
      : null;

    function utilYear(type, year) {
      var u = usage[type] || {};
      return Number(u[year]) || 0;
    }

    function entitlementFor(type) {
      if (core && core.entitlements && core.entitlements[type] != null) {
        var e = core.entitlements[type];
        if (e === 'Unlimited') return 0;
        return Number(e) || 0;
      }
      return 0;
    }

    var cyEnt = entitlementFor('Annual Leave') + entitlementFor('Probation Leave');
    var cyUtil = utilYear('Annual Leave', cy) + utilYear('Casual Leave', cy) +
      utilYear('Examination Leave', cy) + utilYear('Probation Leave', cy);
    var cyBal = cyEnt - cyUtil;

    emp[cy + ' Entitlement'] = cyEnt;
    emp[cy + ' Utilized'] = cyUtil;
    emp[cy + ' Balance'] = cyBal;

    if (carryActive) {
      var pyEnt = entitlementFor('Annual Leave') + entitlementFor('Probation Leave');
      var pyUtil = utilYear('Annual Leave', py) + utilYear('Casual Leave', py) +
        utilYear('Examination Leave', py) + utilYear('Probation Leave', py);
      emp[py + ' Entitlement'] = pyEnt;
      emp[py + ' Utilized'] = pyUtil;
      emp[py + ' Balance'] = pyEnt - pyUtil;
    } else {
      emp[py + ' Entitlement'] = '';
      emp[py + ' Utilized'] = '';
      emp[py + ' Balance'] = '';
    }

    var row = emp._sheetRow;
    if (!row) return;
    [py, cy].forEach(function (y) {
      ['Entitlement', 'Utilized', 'Balance'].forEach(function (suffix) {
        var name = y + ' ' + suffix;
        var c = colIdx[name];
        if (c == null || c < 0) return;
        var val = emp[name];
        if (val === undefined) val = '';
        empSheet.getRange(row, c + 1).setValue(val);
      });
    });
  });
}

function isCarryForwardActive_(policies, today) {
  today = today || new Date();
  var cy = today.getFullYear();
  for (var i = 0; i < (policies || []).length; i++) {
    var pol = policies[i];
    if (String(pol['Leave Type'] || '').trim() !== 'Annual Leave') continue;
    var deadlineRaw = pol['Carry Forward Deadline'];
    if (deadlineRaw === null || deadlineRaw === undefined) continue;
    var s = String(deadlineRaw).trim().toLowerCase();
    if (s === '' || s === 'no' || s === 'none') continue;
    var deadline = null;
    if (deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) {
      deadline = new Date(cy, deadlineRaw.getMonth(), deadlineRaw.getDate());
    } else if (typeof parseDDMMYYYY === 'function') {
      var tmp = parseDDMMYYYY(String(deadlineRaw));
      if (tmp && !isNaN(tmp.getTime())) deadline = new Date(cy, tmp.getMonth(), tmp.getDate());
    }
    if (deadline && today.getTime() <= deadline.getTime()) return true;
  }
  return false;
}

function exportEntitlementAndBalance(buFilter, deptFilter) {
  if (typeof exportEntitlementBalanceCsv === 'function') {
    var csv = exportEntitlementBalanceCsv(buFilter, deptFilter);
    return { success: true, csv: csv, message: 'Ready', fileName: 'Entitlement_Balance.csv' };
  }
  return { success: false, message: 'Export helper missing.' };
}
