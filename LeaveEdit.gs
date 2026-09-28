/**
 * Update leave record — supports leave type change + recalculation.
 * (Overrides older updateLeaveRecord in Code.gs if both present —
 *  remove the Code.gs copy to avoid duplicate function errors.)
 */
function updateLeaveRecord(updateData, userSession) {
  var tblLeave = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  if (!tblLeave) return { success: false, message: 'tblLeave missing.' };
  var data = tblLeave.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var entryCodeIdx = headers.indexOf('Entry Code');
  var target = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][entryCodeIdx]).trim() === updateData.entryCode) {
      target = i + 1;
      break;
    }
  }
  if (target < 0) return { success: false, message: 'Record not found.' };

  var sDate = typeof parseDateOnly_ === 'function'
    ? parseDateOnly_(updateData.startDate)
    : new Date(updateData.startDate);
  var eDate = typeof parseDateOnly_ === 'function'
    ? parseDateOnly_(updateData.endDate)
    : new Date(updateData.endDate);
  if (!sDate || !eDate || isNaN(sDate.getTime()) || isNaN(eDate.getTime())) {
    return { success: false, message: 'Invalid dates.' };
  }

  var typeIdx = headers.indexOf('Leave Type');
  var type = updateData.leaveType
    ? String(updateData.leaveType).trim()
    : String(data[target - 1][typeIdx] || '');

  var noOfDays = Math.round((eDate - sDate) / 86400000) + 1;
  var empId = String(updateData.empId || data[target - 1][headers.indexOf('Emp ID')] || '').trim().toUpperCase();
  var utilized = noOfDays;
  try {
    if (typeof calculateLeaveUtilize === 'function') {
      utilized = calculateLeaveUtilize(empId, sDate, eDate, type);
    }
  } catch (e) {
    utilized = noOfDays;
  }

  var updates = [
    { col: 'Start Date', val: sDate },
    { col: 'End Date', val: eDate },
    { col: 'Leave Reason', val: updateData.leaveReason },
    { col: 'No of Days', val: noOfDays },
    { col: 'Leave Utilized', val: utilized },
    { col: 'Date Modified', val: new Date() },
    { col: 'Modified By', val: userSession && userSession.name ? userSession.name : '' }
  ];
  if (type) updates.push({ col: 'Leave Type', val: type });

  updates.forEach(function (u) {
    var c = headers.indexOf(u.col);
    if (c > -1) {
      var cell = tblLeave.getRange(target, c + 1);
      cell.setValue(u.val);
      if (u.col === 'Start Date' || u.col === 'End Date') {
        try { cell.setNumberFormat('dd-mmm-yyyy'); } catch (fe) {}
      }
    }
  });

  if (typeof invalidateEmpCaches_ === 'function') invalidateEmpCaches_(empId);
  return { success: true, message: 'Record ' + updateData.entryCode + ' updated.' };
}

/**
 * getEmployeeForForm — also resolve by partial name if unique.
 */
function getEmployeeForForm(empId) {
  var search = String(empId || '').trim().toUpperCase();
  var empMap = loadEmployeeMapCached_();
  var emp = empMap[search];
  if (!emp) {
    var needle = String(empId || '').trim().toLowerCase();
    var partials = [];
    Object.keys(empMap).forEach(function (id) {
      if (id === '_headers') return;
      var e = empMap[id];
      var name = String(e['Emp Name'] || '').toLowerCase();
      if (id.indexOf(search) !== -1 || name.indexOf(needle) !== -1) partials.push(id);
    });
    if (partials.length === 1) emp = empMap[partials[0]];
    else if (partials.length > 1) return { error: 'Multiple employees match. Select from suggestions.' };
    else return { error: 'Employee ID not found.' };
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
  Object.keys(balancePayload.balances || {}).forEach(function (type) {
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
