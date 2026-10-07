/**
 * LEAVE BALANCE — daily tblLeave
 * Year columns: Entitlement = Annual + Probation policy ents
 *               Utilized = Annual + Casual + Examination + Probation util for that year
 *               Balance = Entitlement - Utilized
 * Export CSV: emp details through DOJ + 6 year cols if CF active, else 3
 */

function apiGetEmployeeBalance(empId) {
  empId = String(empId || '').trim().toUpperCase();
  if (!empId) return { error: 'Employee Not Found' };
  var pack = loadBalancePackForEmp_(empId);
  if (pack.error) return pack;
  var result = computeBalancesForEmployeeCore_(pack.profile, pack.policies, pack.usage, pack.today);
  return {
    profile: { id: empId, name: pack.profile.empName, bu: pack.profile.bu, dept: pack.profile.department,
      category: pack.profile.category, gender: pack.profile.gender, doj: pack.profile.doj },
    balances: result.balances, entitlements: result.entitlements, usage: result.usageOut,
    detail: result.detail, entitledTypes: result.entitledTypes
  };
}

function loadBalancePackForEmp_(empId) {
  var empMap = (typeof loadEmployeeMapCached_ === 'function') ? loadEmployeeMapCached_() : {};
  var emp = empMap[empId];
  if (!emp) return { error: 'Employee Not Found' };
  var policies = (typeof loadPoliciesCached_ === 'function') ? loadPoliciesCached_() : [];
  var leaveRows = (typeof loadLeaveRowsForEmp_ === 'function') ? loadLeaveRowsForEmp_(empId) : [];
  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;
  var carryGross = getStartingBalanceForEmp_(empId, prevYear);
  var usage = {};
  for (var i = 0; i < leaveRows.length; i++) {
    var lr = leaveRows[i];
    var lt = String(lr['Leave Type'] || '').trim();
    var util = Number(lr['Leave Utilized']) || 0;
    var ey = Number(lr['Entitlement Year']);
    if (!ey) {
      var sd = (lr['Leave Date'] != null && lr['Leave Date'] !== '') ? lr['Leave Date'] : lr['Start Date'];
      if (sd instanceof Date && !isNaN(sd.getTime())) ey = sd.getFullYear();
      else if (typeof dailyDateKey_ === 'function') {
        var dk = dailyDateKey_(sd);
        ey = dk ? Number(dk.substring(0, 4)) : currentYear;
      } else ey = currentYear;
    }
    if (!lt) continue;
    if (!usage[lt]) usage[lt] = {};
    if (!usage[lt][ey]) usage[lt][ey] = 0;
    usage[lt][ey] += util;
  }
  var dojRaw = emp['Date of Join'];
  var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
  return {
    profile: {
      empId: empId, bu: String(emp['Business Unit'] || '').trim(), category: String(emp['Category'] || '').trim(),
      status: String(emp['Status'] || '').trim(), gender: String(emp['Gender'] || '').trim(),
      empName: String(emp['Emp Name'] || '').trim(), department: String(emp['Department'] || '').trim(),
      doj: doj, carryGross: carryGross
    }, policies: policies, usage: usage, today: today
  };
}

function getStartingBalanceForEmp_(empId, year) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sb = ss.getSheetByName('StartingBal');
    if (!sb) return 0;
    var sbData = sb.getDataRange().getValues();
    if (sbData.length < 2) return 0;
    var sbH = sbData[0].map(function (h) { return String(h).trim(); });
    var eIdx = sbH.indexOf('Emp No'); if (eIdx < 0) eIdx = sbH.indexOf('Emp ID');
    var bIdx = sbH.indexOf(year + ' Balance');
    if (eIdx < 0 || bIdx < 0) return 0;
    empId = String(empId).trim().toUpperCase();
    for (var s = 1; s < sbData.length; s++) {
      if (String(sbData[s][eIdx]).trim().toUpperCase() === empId) return Number(sbData[s][bIdx]) || 0;
    }
  } catch (e) {}
  return 0;
}

function batchComputeEmployeeBalances_(employees, policies, leaveUsageByEmp, startingBalMap, today, leaveColMap) {
  leaveColMap = leaveColMap || defaultLeaveColMap_();
  today = today || new Date();
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    var empId = String(emp['Emp ID'] || '').trim().toUpperCase();
    var dojRaw = emp['Date of Join'];
    var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
    var core = computeBalancesForEmployeeCore_({
      empId: empId, bu: String(emp['Business Unit'] || '').trim(), category: String(emp['Category'] || '').trim(),
      status: String(emp['Status'] || '').trim(), gender: String(emp['Gender'] || '').trim(),
      empName: String(emp['Emp Name'] || '').trim(), department: String(emp['Department'] || '').trim(),
      doj: doj, carryGross: Number(startingBalMap[empId]) || 0
    }, policies, leaveUsageByEmp[empId] || {}, today);
    Object.keys(leaveColMap).forEach(function (col) {
      var bal = core.balances[leaveColMap[col]];
      if (bal === undefined || bal === null) emp[col] = '';
      else if (bal === 'Unlimited') emp[col] = 'Unlimited';
      else emp[col] = Number(bal);
    });
  }
}

function defaultLeaveColMap_() {
  return { 'Annual': 'Annual Leave', 'Casual': 'Casual Leave', 'Compassionate': 'Compassionate Leave',
    'Examination': 'Examination Leave', 'Study': 'Study Leave', 'Maternity': 'Maternity Leave', 'Probation': 'Probation Leave' };
}

function yearChargeLeaveTypes_() {
  return ['Annual Leave', 'Casual Leave', 'Examination Leave', 'Probation Leave'];
}

function yearColNames_(year) {
  return { ent: year + ' Entitlement', util: year + ' Utilized', bal: year + ' Balance' };
}

function computeYearRollup_(typeMeta, usage, year) {
  function utilOf(lt, y) { var u = usage[lt] || {}; return Number(u[y]) || 0; }
  var annualEnt = 0, probationEnt = 0;
  if (typeMeta['Annual Leave'] && !typeMeta['Annual Leave'].isUnlimited)
    annualEnt = Number(typeMeta['Annual Leave'].entitlement) || 0;
  if (typeMeta['Probation Leave'] && !typeMeta['Probation Leave'].isUnlimited)
    probationEnt = Number(typeMeta['Probation Leave'].entitlement) || 0;
  var entitlement = annualEnt + probationEnt;
  var utilized = 0;
  var types = yearChargeLeaveTypes_();
  for (var i = 0; i < types.length; i++) utilized += utilOf(types[i], year);
  return { entitlement: entitlement, utilized: utilized, balance: entitlement - utilized };
}

function computeBalancesForEmployeeCore_(profile, policies, usage, today) {
  today = today || new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;
  var bu = profile.bu, catFull = profile.category, empStatus = profile.status, gender = profile.gender;
  var doj = profile.doj instanceof Date ? profile.doj : parseDDMMYYYY(String(profile.doj || ''));
  var catInitials = mapCategoryToInitials(catFull);
  var carryGross = Number(profile.carryGross) || 0;
  var typeMeta = {};
  for (var p = 0; p < policies.length; p++) {
    var pol = policies[p];
    var lType = String(pol['Leave Type'] || '').trim();
    if (!lType) continue;
    var scoreBU = checkMatch(bu.toUpperCase(), String(pol['Business Unit'] || '').trim().toUpperCase(), 100);
    var scoreCat = checkMatch(catInitials, String(pol['Category'] || '').trim(), 10);
    if (scoreCat === -1) scoreCat = checkMatch(catFull, String(pol['Category'] || '').trim(), 10);
    var scoreStatus = checkMatch(empStatus, String(pol['Status'] || '').trim(), 5);
    var scoreGender = checkMatch(gender, String(pol['Gender'] || '').trim(), 1);
    if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) continue;
    if (!evaluateLifecycle(doj, String(pol['Who is entitled (Lifecycle)'] || '').trim(), today)) continue;
    var totalScore = scoreBU + scoreCat + scoreStatus + scoreGender;
    var entRaw = String(pol['Annual Entitlements'] || '').trim();
    var isUnlimited = (entRaw === '' || entRaw.toLowerCase() === 'unlimited');
    var entVal = isUnlimited ? 'Unlimited' : (Number(entRaw) || 0);
    var show = String(pol['Balance Page Show'] || '').trim().toLowerCase();
    var showYes = (show === 'yes' || show === 'y' || show === 'true');
    var deductFrom = String(pol['Deduct from'] || '').trim();
    if (deductFrom.toLowerCase() === 'none' || deductFrom.toLowerCase() === 'null') deductFrom = '';
    var deadlineRaw = pol['Carry Forward Deadline'];
    var deadline = null;
    if (deadlineRaw !== null && deadlineRaw !== undefined && String(deadlineRaw).trim() !== '' && String(deadlineRaw).trim().toLowerCase() !== 'no') {
      if (deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) deadline = new Date(currentYear, deadlineRaw.getMonth(), deadlineRaw.getDate());
      else {
        var tmp = parseDDMMYYYY(String(deadlineRaw));
        if (!isNaN(tmp.getTime())) deadline = new Date(currentYear, tmp.getMonth(), tmp.getDate());
      }
    }
    if (!typeMeta[lType] || totalScore > typeMeta[lType].score) {
      typeMeta[lType] = { score: totalScore, entitlement: entVal, isUnlimited: isUnlimited, show: showYes, deductFrom: deductFrom, deadline: deadline };
    }
  }
  var attachedToAnnual = [];
  Object.keys(typeMeta).forEach(function (t) {
    var df = String(typeMeta[t].deductFrom || '').trim().toLowerCase();
    if (df === 'annual leave' || df === 'annual') attachedToAnnual.push(t);
  });
  if (attachedToAnnual.indexOf('Probation Leave') < 0) attachedToAnnual.push('Probation Leave');
  function utilOf(leaveType, year) { var u = usage[leaveType] || {}; return Number(u[year]) || 0; }
  function annualChargeForYear_(year) {
    var total = utilOf('Annual Leave', year);
    for (var a = 0; a < attachedToAnnual.length; a++) total += utilOf(attachedToAnnual[a], year);
    return total;
  }
  var annualMeta = typeMeta['Annual Leave'] || null;
  var carryActive = false, carryDeadline = null;
  if (annualMeta && annualMeta.deadline) {
    carryDeadline = annualMeta.deadline;
    var todayDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var dlDay = new Date(carryDeadline.getFullYear(), carryDeadline.getMonth(), carryDeadline.getDate());
    carryActive = todayDay.getTime() <= dlDay.getTime();
  }
  var prevYearCharge = annualChargeForYear_(prevYear);
  var prevAvail = carryActive ? Math.max(0, carryGross - prevYearCharge) : 0;
  var balances = {}, entitlements = {}, usageOut = {}, detail = {};
  Object.keys(typeMeta).forEach(function (t) {
    var meta = typeMeta[t];
    entitlements[t] = meta.entitlement;
    var cy = utilOf(t, currentYear), py = utilOf(t, prevYear);
    usageOut[t] = cy + (carryActive ? py : 0);
    if (meta.isUnlimited) {
      balances[t] = 'Unlimited';
      detail[t] = { prevYearBalance: 0, prevYearUtilized: py, thisYearEntitlement: 'Unlimited', thisYearUtilized: cy,
        thisYearBalance: 'Unlimited', currentBalance: 'Unlimited', carryExpired: !carryActive,
        carryDeadline: carryDeadline ? leaveDateKeySafe_(carryDeadline) : null, deductFrom: meta.deductFrom, show: meta.show };
      return;
    }
    var ent = Number(meta.entitlement) || 0, thisYearBal = 0, prevBalShown = 0;
    if (t === 'Annual Leave') {
      thisYearBal = Math.max(0, ent - annualChargeForYear_(currentYear));
      prevBalShown = prevAvail;
      balances[t] = thisYearBal + prevBalShown;
    } else {
      thisYearBal = Math.max(0, ent - cy);
      balances[t] = thisYearBal;
    }
    detail[t] = { prevYearBalance: prevBalShown, prevYearUtilized: py,
      prevYearChargeAgainstAnnual: (t === 'Annual Leave') ? prevYearCharge : null,
      thisYearEntitlement: ent, thisYearUtilized: cy, thisYearBalance: thisYearBal, currentBalance: balances[t],
      carryExpired: !carryActive, carryDeadline: carryDeadline ? leaveDateKeySafe_(carryDeadline) : null,
      deductFrom: meta.deductFrom, show: meta.show };
  });
  var annualBal = balances['Annual Leave'];
  if (typeof annualBal === 'number') {
    attachedToAnnual.forEach(function (t) {
      if (t === 'Probation Leave' || typeof balances[t] !== 'number') return;
      var capped = Math.min(balances[t], annualBal);
      balances[t] = capped;
      if (detail[t]) { detail[t].currentBalance = capped; detail[t].thisYearBalance = capped; }
    });
  }
  return { balances: balances, entitlements: entitlements, usageOut: usageOut, detail: detail,
    entitledTypes: Object.keys(typeMeta), carryActive: carryActive, prevAvail: prevAvail, typeMeta: typeMeta };
}

function leaveDateKeySafe_(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  catch (e) {
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
}

function updateAllEmployeeLeaveBalances() {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  var leaveSheet = ss.getSheetByName('tblLeave');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var sbSheet = ss.getSheetByName('StartingBal');
  if (!empSheet) return { success: false, message: 'tblEmployee missing.' };
  if (!policySheet) return { success: false, message: 'Sys_LeavePolicies missing.' };
  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;
  var policies = [];
  var pData = policySheet.getDataRange().getValues();
  if (pData.length > 1) {
    var pH = pData[0].map(function (h) { return String(h).trim(); });
    for (var i = 1; i < pData.length; i++) {
      var o = {}; for (var c = 0; c < pH.length; c++) o[pH[c]] = pData[i][c];
      policies.push(o);
    }
  }
  var startingBalMap = {};
  if (sbSheet) {
    var sbData = sbSheet.getDataRange().getValues();
    if (sbData.length > 1) {
      var sbH = sbData[0].map(function (h) { return String(h).trim(); });
      var eIdx = sbH.indexOf('Emp No'); if (eIdx < 0) eIdx = sbH.indexOf('Emp ID');
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
      var liEmp = lH.indexOf('Emp ID'), liType = lH.indexOf('Leave Type');
      var liUtil = lH.indexOf('Leave Utilized'), liYear = lH.indexOf('Entitlement Year');
      var liStart = lH.indexOf('Start Date'), liDate = lH.indexOf('Leave Date');
      for (var r = 1; r < lData.length; r++) {
        var eid = liEmp >= 0 ? String(lData[r][liEmp] || '').trim().toUpperCase() : '';
        var lt = liType >= 0 ? String(lData[r][liType] || '').trim() : '';
        var util = liUtil >= 0 ? Number(lData[r][liUtil]) || 0 : 0;
        var ey = liYear >= 0 ? Number(lData[r][liYear]) : 0;
        if (!ey) {
          var dCell = liDate >= 0 ? lData[r][liDate] : (liStart >= 0 ? lData[r][liStart] : null);
          if (dCell instanceof Date && !isNaN(dCell.getTime())) {
            ey = dCell.getFullYear();
            if (typeof dailyDateKey_ === 'function') { var dks = dailyDateKey_(dCell); if (dks) ey = Number(dks.substring(0, 4)); }
          } else if (typeof dailyDateKey_ === 'function') {
            var dks2 = dailyDateKey_(dCell); if (dks2) ey = Number(dks2.substring(0, 4));
          }
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
  if (eData.length < 2) return { success: true, message: 'No employees.', updated: 0 };
  var eH = eData[0].map(function (h) { return String(h).trim(); });
  var leaveColMap = defaultLeaveColMap_();
  var colIdx = {};
  Object.keys(leaveColMap).forEach(function (col) { colIdx[col] = eH.indexOf(col); });
  var prevNames = yearColNames_(prevYear), currNames = yearColNames_(currentYear);
  var yearCols = [prevNames.ent, prevNames.util, prevNames.bal, currNames.ent, currNames.util, currNames.bal];
  yearCols.forEach(function (c) { colIdx[c] = eH.indexOf(c); });
  var employees = [];
  for (var er = 1; er < eData.length; er++) {
    var rowObj = {}; for (var ec = 0; ec < eH.length; ec++) rowObj[eH[ec]] = eData[er][ec];
    rowObj._sheetRow = er + 1; employees.push(rowObj);
  }
  batchComputeEmployeeBalances_(employees, policies, leaveUsageByEmp, startingBalMap, today, leaveColMap);
  for (var ei = 0; ei < employees.length; ei++) {
    var emp = employees[ei];
    var empId = String(emp['Emp ID'] || '').trim().toUpperCase();
    var usage = leaveUsageByEmp[empId] || {};
    var dojRaw = emp['Date of Join'];
    var doj = (dojRaw instanceof Date) ? dojRaw : parseDDMMYYYY(String(dojRaw || ''));
    var core = computeBalancesForEmployeeCore_({
      empId: empId, bu: String(emp['Business Unit'] || '').trim(), category: String(emp['Category'] || '').trim(),
      status: String(emp['Status'] || '').trim(), gender: String(emp['Gender'] || '').trim(),
      empName: String(emp['Emp Name'] || '').trim(), department: String(emp['Department'] || '').trim(),
      doj: doj, carryGross: Number(startingBalMap[empId]) || 0
    }, policies, usage, today);
    var typeMeta = core.typeMeta || {};
    var prevR = computeYearRollup_(typeMeta, usage, prevYear);
    var currR = computeYearRollup_(typeMeta, usage, currentYear);
    emp[prevNames.ent] = prevR.entitlement; emp[prevNames.util] = prevR.utilized; emp[prevNames.bal] = prevR.balance;
    emp[currNames.ent] = currR.entitlement; emp[currNames.util] = currR.utilized; emp[currNames.bal] = currR.balance;
  }
  var updated = 0;
  var writeCols = Object.keys(leaveColMap).concat(yearCols);
  employees.forEach(function (emp) {
    writeCols.forEach(function (col) {
      var ci = colIdx[col]; if (ci === undefined || ci < 0) return;
      var val = emp[col]; if (val === undefined) val = '';
      empSheet.getRange(emp._sheetRow, ci + 1).setValue(val);
    });
    updated++;
  });
  SpreadsheetApp.flush();
  if (typeof cacheClearAll_ === 'function') cacheClearAll_();
  var ms = new Date().getTime() - started;
  var msg = 'Updated leave balances + year rollups for ' + updated + ' employee(s) in ' + ms + ' ms.';
  Logger.log(msg);
  return { success: true, message: msg, updated: updated, elapsedMs: ms };
}

function exportAnnualCasualBalances() {
  var prep = updateAllEmployeeLeaveBalances();
  if (prep && prep.success === false) return prep;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!empSheet) return { success: false, message: 'tblEmployee missing.' };
  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;
  var carryActive = true, deadlineStr = '';
  try {
    var policySheet = ss.getSheetByName('Sys_LeavePolicies');
    if (policySheet) {
      var pData = policySheet.getDataRange().getValues();
      var pH = pData[0].map(function (h) { return String(h).trim(); });
      var iType = pH.indexOf('Leave Type'), iDl = pH.indexOf('Carry Forward Deadline');
      for (var i = 1; i < pData.length; i++) {
        if (String(pData[i][iType] || '').trim() !== 'Annual Leave') continue;
        var deadlineRaw = pData[i][iDl];
        if (deadlineRaw === null || deadlineRaw === undefined || String(deadlineRaw).trim() === '' || String(deadlineRaw).trim().toLowerCase() === 'no') continue;
        var deadline;
        if (deadlineRaw instanceof Date && !isNaN(deadlineRaw.getTime())) deadline = new Date(currentYear, deadlineRaw.getMonth(), deadlineRaw.getDate());
        else {
          var tmp = parseDDMMYYYY(String(deadlineRaw));
          if (isNaN(tmp.getTime())) continue;
          deadline = new Date(currentYear, tmp.getMonth(), tmp.getDate());
        }
        var todayDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        var dlDay = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
        carryActive = todayDay.getTime() <= dlDay.getTime();
        deadlineStr = leaveDateKeySafe_(dlDay);
        break;
      }
    }
  } catch (e) { Logger.log('CF detect: ' + e.message); }
  var eData = empSheet.getDataRange().getValues();
  if (eData.length < 2) return { success: false, message: 'No employees.' };
  var eH = eData[0].map(function (h) { return String(h).trim(); });
  var baseCols = ['Emp ID', 'Emp Name', 'Business Unit', 'Department', 'Category', 'Gender', 'Status', 'Date of Join'];
  function findCol_(name) {
    var ix = eH.indexOf(name);
    if (ix >= 0) return ix;
    if (name === 'Date of Join') {
      ix = eH.indexOf('Date of Joining'); if (ix < 0) ix = eH.indexOf('DOJ');
    }
    return ix;
  }
  var prevN = yearColNames_(prevYear), currN = yearColNames_(currentYear);
  var yearCols = carryActive
    ? [prevN.ent, prevN.util, prevN.bal, currN.ent, currN.util, currN.bal]
    : [currN.ent, currN.util, currN.bal];
  var headers = baseCols.concat(yearCols);
  var lines = [headers.map(csvEscape_).join(',')];
  for (var r = 1; r < eData.length; r++) {
    var row = [];
    for (var c = 0; c < baseCols.length; c++) {
      var ix = findCol_(baseCols[c]);
      var v = ix >= 0 ? eData[r][ix] : '';
      if (v instanceof Date) v = leaveDateKeySafe_(v);
      row.push(csvEscape_(v));
    }
    for (var y = 0; y < yearCols.length; y++) {
      var yx = eH.indexOf(yearCols[y]);
      row.push(csvEscape_(yx >= 0 ? eData[r][yx] : ''));
    }
    lines.push(row.join(','));
  }
  var csv = lines.join('\n');
  var stamp = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  var filename = 'Leave_Year_Balances_' + stamp + '.csv';
  var url = '';
  try { var file = DriveApp.createFile(filename, csv, MimeType.CSV); url = file.getDownloadUrl(); }
  catch (e2) { Logger.log('Drive CSV create failed: ' + e2.message); }
  try {
    var tabName = 'Export_Leave_Balances';
    var existing = ss.getSheetByName(tabName); if (existing) ss.deleteSheet(existing);
    var sh = ss.insertSheet(tabName);
    var values = [headers];
    for (var r2 = 1; r2 < eData.length; r2++) {
      var row2 = [];
      for (var c2 = 0; c2 < baseCols.length; c2++) {
        var ix2 = findCol_(baseCols[c2]);
        var v2 = ix2 >= 0 ? eData[r2][ix2] : '';
        if (v2 instanceof Date) v2 = leaveDateKeySafe_(v2);
        row2.push(v2);
      }
      for (var y2 = 0; y2 < yearCols.length; y2++) {
        var yx2 = eH.indexOf(yearCols[y2]);
        row2.push(yx2 >= 0 ? eData[r2][yx2] : '');
      }
      values.push(row2);
    }
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    sh.setFrozenRows(1);
  } catch (e3) { Logger.log('sheet export tab: ' + e3.message); }
  var msg = 'CSV ' + filename + ' | CF active=' + carryActive + (deadlineStr ? (' deadline=' + deadlineStr) : '') + ' | cols=' + yearCols.join(', ');
  Logger.log(msg);
  return { success: true, message: msg, filename: filename, url: url, csv: csv, carryActive: carryActive, yearColumns: yearCols, rowCount: eData.length - 1 };
}

function csvEscape_(val) {
  if (val === null || val === undefined) return '';
  var s = String(val);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
