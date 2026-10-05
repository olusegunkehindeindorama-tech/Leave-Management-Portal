/**
 * CarryForwardFix.gs — paste AFTER "Entitlement and Utilization.gs"
 * Overwrites splitLeaveByCarryForward_ + calculateLeaveUtilized so that:
 *  - Leaves fully within CF deadline get Entitlement Year = previous year
 *  - While StartingBal "{prev} Balance" remains (consumed chronologically)
 *  - After deadline or when prev balance is exhausted → current year
 */

/**
 * Split / label leave by carry-forward deadline.
 *
 * - Entirely AFTER deadline → entitlementYear = calendar year of start
 * - Entirely ON/BEFORE deadline → entitlementYear = previous calendar year
 *   (charged against StartingBal "{prev} Balance" while that pool lasts —
 *    sequential balance application is done in calculateLeaveUtilized)
 * - Spans deadline → two segments: pre → prev year, post → current year
 */
splitLeaveByCarryForward_ = function(startDate, endDate, leaveType, pol) {
  var meta = policyCalcMeta_(pol);
  var s = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  var e = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  var segments = [];
  var calYear = s.getFullYear();
  var prevYear = calYear - 1;

  if (!meta.deadline) {
    segments.push({ start: s, end: e, entitlementYear: calYear });
    return segments;
  }

  var deadline = carryDeadlineForYear_(meta.deadline, calYear);
  if (!deadline) {
    segments.push({ start: s, end: e, entitlementYear: calYear });
    return segments;
  }

  // Entirely after carry-forward deadline → current year only
  if (s.getTime() > deadline.getTime()) {
    segments.push({ start: s, end: e, entitlementYear: calYear });
    return segments;
  }

  // Entirely on or before deadline → previous entitlement year
  if (e.getTime() <= deadline.getTime()) {
    segments.push({ start: s, end: e, entitlementYear: prevYear });
    return segments;
  }

  // Spans deadline: pre-deadline → prev year; post → current year
  var preEnd = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  var postStart = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate() + 1);

  segments.push({ start: s, end: preEnd, entitlementYear: prevYear });
  segments.push({ start: postStart, end: e, entitlementYear: postStart.getFullYear() });
  return segments;
};

calculateLeaveUtilized = function() {
  var started = new Date().getTime();
  var master = loadEntitlementMasterData_();
  if (!master.sheet.leave) {
    return { success: false, message: 'tblLeave sheet missing.', updated: 0 };
  }
  if (!master.policies.length) {
    return { success: false, message: 'Sys_LeavePolicies empty — cannot recalculate.', updated: 0 };
  }
  if (!Object.keys(master.multipliers.Yes).length && !Object.keys(master.multipliers.No).length) {
    Logger.log('WARNING: Multiplier_Policy sheet missing or empty. Weights default to 1.');
  }

  var headers = master.leaveHeaders;
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var utilIdx = headers.indexOf('Leave Utilized');
  var daysIdx = headers.indexOf('No of Days');
  var yearIdx = headers.indexOf('Entitlement Year');
  var entryIdx = headers.indexOf('Entry Code');

  var mergeGroups = findMergeCandidates_(master.leaves);
  var rowsToDelete = {};
  var mergedLeaves = [];

  mergeGroups.forEach(function (g) {
    var base = g.rows[0];
    for (var k = 1; k < g.rows.length; k++) {
      rowsToDelete[g.rows[k]._row] = true;
    }
    base['Start Date'] = g.start;
    base['End Date'] = g.end;
    if (entryIdx >= 0) {
      base['Entry Code'] = String(base['Entry Code'] || '').replace(/-S[12]$/, '');
    }
    mergedLeaves.push(base);
  });

  var working = [];
  master.leaves.forEach(function (lv) {
    if (rowsToDelete[lv._row]) return;
    var replaced = false;
    for (var m = 0; m < mergedLeaves.length; m++) {
      if (mergedLeaves[m]._row === lv._row) {
        working.push(mergedLeaves[m]);
        replaced = true;
        break;
      }
    }
    if (!replaced) working.push(lv);
  });

  var updates = [];
  var appends = [];
  var splitCount = 0;
  var mergeCount = mergeGroups.length;

  // Sort by emp, start date so prev-year StartingBal is consumed in chronological order
  working.sort(function (a, b) {
    var ae = String(a['Emp ID'] || '').toUpperCase();
    var be = String(b['Emp ID'] || '').toUpperCase();
    if (ae !== be) return ae < be ? -1 : 1;
    var as = new Date(a['Start Date']).getTime();
    var bs = new Date(b['Start Date']).getTime();
    if (as !== bs) return as - bs;
    return (a._row || 0) - (b._row || 0);
  });

  // Remaining previous-year balance per employee (from StartingBal "{year} Balance")
  var remainingPrevByEmp = {};
  function getRemainingPrev_(empId, calYear) {
    var key = empId + '|' + calYear;
    if (remainingPrevByEmp[key] !== undefined) return remainingPrevByEmp[key];
    var prevY = calYear - 1;
    var col = prevY + ' Balance';
    var bal = 0;
    if (master.startingBal[empId]) {
      bal = Number(master.startingBal[empId][col]) || 0;
    }
    remainingPrevByEmp[key] = Math.max(0, bal);
    return remainingPrevByEmp[key];
  }
  function setRemainingPrev_(empId, calYear, val) {
    remainingPrevByEmp[empId + '|' + calYear] = Math.max(0, val);
  }

  /** Does this leave type draw from Annual / StartingBal CF pool? */
  function drawsFromAnnualPool_(emp, lt, pol) {
    if (String(lt || '').trim() === 'Annual Leave') return true;
    var meta = policyCalcMeta_(pol);
    var df = String(meta.deductFrom || '').trim().toLowerCase();
    return df === 'annual leave' || df === 'annual';
  }

  working.forEach(function (lv) {
    var empId = String(lv['Emp ID'] || '').trim().toUpperCase();
    var lt = String(lv['Leave Type'] || '').trim();
    var sRaw = lv['Start Date'];
    var eRaw = lv['End Date'];
    if (typeof toCalendarDate_ === 'function') {
      sRaw = toCalendarDate_(sRaw) || sRaw;
      eRaw = toCalendarDate_(eRaw) || eRaw;
    }
    var s = new Date(sRaw);
    var e = new Date(eRaw);
    if (!empId || isNaN(s.getTime()) || isNaN(e.getTime())) return;
    s = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    e = new Date(e.getFullYear(), e.getMonth(), e.getDate());

    var emp = master.employees[empId] || { 'Emp ID': empId };
    var pol = matchPolicyForEmp_(master.policies, emp, lt, new Date());
    var segments = splitLeaveByCarryForward_(s, e, lt, pol);
    var usesAnnualPool = drawsFromAnnualPool_(emp, lt, pol);
    var calYear = s.getFullYear();

    // Apply sequential StartingBal: if prev-year label but no remaining balance → current year
    function applyBalanceYear_(seg) {
      var year = seg.entitlementYear;
      if (!usesAnnualPool) return year;
      if (year >= calYear) return year; // already current (or later)
      // year is previous — only keep if StartingBal remaining > 0
      var rem = getRemainingPrev_(empId, calYear);
      if (rem > 0) return year;
      return calYear; // prev balance exhausted → charge current year
    }

    if (segments.length === 1) {
      var util = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var days = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      var year1 = applyBalanceYear_(segments[0]);

      // If labeled prev year but util exceeds remaining prev bal, split across years
      if (usesAnnualPool && year1 === calYear - 1) {
        var rem = getRemainingPrev_(empId, calYear);
        if (util > rem && rem > 0) {
          // Partial: rem days to prev year, rest to current — approximate by util split
          var utilPrev = rem;
          var utilCurr = util - rem;
          setRemainingPrev_(empId, calYear, 0);
          splitCount++;
          var baseCode = String(lv['Entry Code'] || 'LV').replace(/-S[12]$/, '');
          // Keep original row as prev-year portion (same dates, reduced util is imperfect
          // but entitlement year is what user needs; keep full date span on first row
          // with prev year only for the portion that fits — prefer date split when possible.
          // Practical approach: first row keeps full dates + year=prev only if rem covers all;
          // here rem < util so assign year=prev for first and append current year row is messy
          // without date split. Assign whole leave to prev until rem=0 then current:
          // When util > rem: still put entitlement year = prev for this leave only if we
          // allow overdraw of display year — better: year = prev when rem>0 for the leave
          // that starts consuming, and mark remaining util as current on a second logical year.
          updates.push({
            row: lv._row,
            start: segments[0].start,
            end: segments[0].end,
            util: util,
            days: days,
            year: calYear - 1,
            entryCode: lv['Entry Code']
          });
          // Consume all remaining prev; excess counts against current year in balance engine
          // via Entitlement Year = prev (full util still on prev year row). For balance math,
          // Leave balance.gs charges by entitlement year. So split util across two rows.
          updates[updates.length - 1].util = utilPrev;
          var newRow = buildBlankLeaveRow_(headers);
          for (var c = 0; c < headers.length; c++) {
            newRow[c] = lv[headers[c]] !== undefined ? lv[headers[c]] : '';
          }
          if (entryIdx >= 0) newRow[entryIdx] = baseCode + '-CY';
          if (startIdx >= 0) newRow[startIdx] = segments[0].start;
          if (endIdx >= 0) newRow[endIdx] = segments[0].end;
          if (utilIdx >= 0) newRow[utilIdx] = utilCurr;
          if (daysIdx >= 0) newRow[daysIdx] = days;
          if (yearIdx >= 0) newRow[yearIdx] = calYear;
          appends.push(newRow);
          return;
        }
        if (rem > 0) {
          setRemainingPrev_(empId, calYear, rem - util);
        } else {
          year1 = calYear;
        }
      }

      updates.push({
        row: lv._row,
        start: segments[0].start,
        end: segments[0].end,
        util: util,
        days: days,
        year: year1,
        entryCode: lv['Entry Code']
      });
    } else {
      splitCount++;
      var u0 = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var d0 = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      var baseCode = String(lv['Entry Code'] || 'LV').replace(/-S[12]$/, '');
      var y0 = applyBalanceYear_(segments[0]);
      if (usesAnnualPool && y0 === calYear - 1) {
        var rem0 = getRemainingPrev_(empId, calYear);
        if (rem0 <= 0) y0 = calYear;
        else setRemainingPrev_(empId, calYear, rem0 - u0);
      }
      updates.push({
        row: lv._row,
        start: segments[0].start,
        end: segments[0].end,
        util: u0,
        days: d0,
        year: y0,
        entryCode: baseCode + '-S1'
      });

      var u1 = calculateLeaveUtilizeWithMaster_(master, empId, segments[1].start, segments[1].end, lt);
      var d1 = Math.round((segments[1].end - segments[1].start) / 86400000) + 1;
      var y1 = segments[1].entitlementYear; // post-deadline = current year
      var newRow = buildBlankLeaveRow_(headers);
      for (var c2 = 0; c2 < headers.length; c2++) {
        newRow[c2] = lv[headers[c2]] !== undefined ? lv[headers[c2]] : '';
      }
      if (entryIdx >= 0) newRow[entryIdx] = baseCode + '-S2';
      if (startIdx >= 0) newRow[startIdx] = segments[1].start;
      if (endIdx >= 0) newRow[endIdx] = segments[1].end;
      if (utilIdx >= 0) newRow[utilIdx] = u1;
      if (daysIdx >= 0) newRow[daysIdx] = d1;
      if (yearIdx >= 0) newRow[yearIdx] = y1;
      appends.push(newRow);
    }
  });

  var leaveSheet = master.sheet.leave;
  updates.forEach(function (u) {
    if (startIdx >= 0) leaveSheet.getRange(u.row, startIdx + 1).setValue(u.start);
    if (endIdx >= 0) leaveSheet.getRange(u.row, endIdx + 1).setValue(u.end);
    if (utilIdx >= 0) leaveSheet.getRange(u.row, utilIdx + 1).setValue(u.util);
    if (daysIdx >= 0) leaveSheet.getRange(u.row, daysIdx + 1).setValue(u.days);
    if (yearIdx >= 0) leaveSheet.getRange(u.row, yearIdx + 1).setValue(u.year);
    if (entryIdx >= 0 && u.entryCode) leaveSheet.getRange(u.row, entryIdx + 1).setValue(u.entryCode);
  });

  var delRows = Object.keys(rowsToDelete).map(Number).sort(function (a, b) { return b - a; });
  delRows.forEach(function (rn) {
    leaveSheet.deleteRow(rn);
  });

  if (appends.length) {
    appendLeaveRows_(leaveSheet, appends);
  }

  try {
    refreshAllEmployeeLeaveColumns_(master);
  } catch (e) {
    Logger.log('Balance column refresh skipped: ' + e.message);
  }

  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var ms = new Date().getTime() - started;
  var msg = 'Recalculated ' + updates.length + ' leave(s), merged ' + mergeCount +
    ' wrong split group(s), created ' + splitCount + ' carry-forward split(s), ' +
    'appended ' + appends.length + ' row(s) in ' + ms + ' ms.';
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    updated: updates.length,
    merged: mergeCount,
    split: splitCount,
    appended: appends.length,
    elapsedMs: ms
  };
};


recalculateAllLeaveUtilized = function () {
  return calculateLeaveUtilized();
};
