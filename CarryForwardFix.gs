/**
 * CarryForwardFix.gs — paste AFTER "Entitlement and Utilization.gs"
 * Overwrites splitLeaveByCarryForward_ + calculateLeaveUtilized so that:
 *  - Leaves fully within CF deadline get Entitlement Year = previous year
 *  - While StartingBal "{prev} Balance" remains (consumed chronologically)
 *  - After deadline or when prev balance is exhausted → current year
 *  - Exact same Emp+Start+End rows (e.g. BP-861-a + BP-861-a-CY) are merged
 *    before re-split so duplicate CY rows cannot stack.
 */

/**
 * Merge candidates:
 *  A) Exact same Emp + Leave Type + calendar Start + End (collapses duplicate -CY rows)
 *  B) Adjacent ranges with -S1/-S2 sibling codes (original behaviour)
 * Entry suffixes -S1, -S2, -CY are stripped when merging so recalc can re-split cleanly.
 */
findMergeCandidates_ = function (leaves) {
  function calDay_(d) {
    if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  function dayKey_(d) {
    var c = calDay_(d);
    if (!c) return '';
    return c.getFullYear() + '-' + ('0' + (c.getMonth() + 1)).slice(-2) + '-' + ('0' + c.getDate()).slice(-2);
  }
  function baseCode_(code) {
    var c = String(code || '').trim();
    c = c.replace(/-CY$/i, '');
    c = c.replace(/-S[12]$/i, '');
    return c;
  }

  var byExact = {};
  for (var i = 0; i < leaves.length; i++) {
    var row = leaves[i];
    var emp = String(row['Emp ID'] || '').trim().toUpperCase();
    var lt = String(row['Leave Type'] || '');
    var sk = dayKey_(row['Start Date']);
    var ek = dayKey_(row['End Date']);
    if (!emp || !sk || !ek) continue;
    var key = emp + '|' + lt + '|' + sk + '|' + ek;
    if (!byExact[key]) byExact[key] = [];
    byExact[key].push(row);
  }
  var groups = [];
  var consumed = {};
  Object.keys(byExact).forEach(function (k) {
    var rows = byExact[k];
    if (rows.length < 2) return;
    rows.sort(function (a, b) {
      var ac = String(a['Entry Code'] || '');
      var bc = String(b['Entry Code'] || '');
      var aCy = /-CY$/i.test(ac) ? 1 : 0;
      var bCy = /-CY$/i.test(bc) ? 1 : 0;
      if (aCy !== bCy) return aCy - bCy;
      return (a._row || 0) - (b._row || 0);
    });
    var s = calDay_(rows[0]['Start Date']);
    var e = calDay_(rows[0]['End Date']);
    groups.push({ rows: rows, start: s, end: e });
    rows.forEach(function (r) { consumed[r._row] = true; });
  });

  var sorted = leaves.filter(function (r) { return !consumed[r._row]; }).slice().sort(function (a, b) {
    var ae = String(a['Emp ID'] || '').toUpperCase();
    var be = String(b['Emp ID'] || '').toUpperCase();
    if (ae !== be) return ae < be ? -1 : 1;
    var at = String(a['Leave Type'] || '');
    var bt = String(b['Leave Type'] || '');
    if (at !== bt) return at < bt ? -1 : 1;
    return new Date(a['Start Date']) - new Date(b['Start Date']);
  });

  var cur = null;
  for (var j = 0; j < sorted.length; j++) {
    var r2 = sorted[j];
    var s2 = calDay_(r2['Start Date']);
    var e2 = calDay_(r2['End Date']);
    if (!s2 || !e2) continue;

    if (!cur) {
      cur = { rows: [r2], start: s2, end: e2 };
      continue;
    }

    var sameEmp = String(cur.rows[0]['Emp ID']).toUpperCase() === String(r2['Emp ID']).toUpperCase();
    var sameType = String(cur.rows[0]['Leave Type']) === String(r2['Leave Type']);
    var nextDay = new Date(cur.end.getFullYear(), cur.end.getMonth(), cur.end.getDate() + 1);
    var adjacent = nextDay.getTime() === s2.getTime();
    var reasonA = String(cur.rows[0]['Leave Reason'] || '');
    var reasonB = String(r2['Leave Reason'] || '');
    var sameReason = reasonA === reasonB;
    var codeA = String(cur.rows[0]['Entry Code'] || '');
    var codeB = String(r2['Entry Code'] || '');
    var splitSibling = (
      /-S[12]$/i.test(codeA) || /-S[12]$/i.test(codeB) ||
      /-CY$/i.test(codeA) || /-CY$/i.test(codeB) ||
      (baseCode_(codeA) === baseCode_(codeB) && baseCode_(codeA) !== '')
    );

    if (sameEmp && sameType && adjacent && (sameReason || splitSibling)) {
      cur.rows.push(r2);
      cur.end = e2;
    } else {
      if (cur.rows.length > 1) groups.push(cur);
      cur = { rows: [r2], start: s2, end: e2 };
    }
  }
  if (cur && cur.rows.length > 1) groups.push(cur);
  return groups;
};

splitLeaveByCarryForward_ = function (startDate, endDate, leaveType, pol) {
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

  if (s.getTime() > deadline.getTime()) {
    segments.push({ start: s, end: e, entitlementYear: calYear });
    return segments;
  }

  if (e.getTime() <= deadline.getTime()) {
    segments.push({ start: s, end: e, entitlementYear: prevYear });
    return segments;
  }

  var preEnd = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  var postStart = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate() + 1);
  segments.push({ start: s, end: preEnd, entitlementYear: prevYear });
  segments.push({ start: postStart, end: e, entitlementYear: postStart.getFullYear() });
  return segments;
};

calculateLeaveUtilized = function () {
  Logger.log('=== RECALC START (calculateLeaveUtilized) ===');
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
  Logger.log('Merge groups found: ' + mergeGroups.length);
  var rowsToDelete = {};
  var mergedLeaves = [];

  mergeGroups.forEach(function (g) {
    var base = g.rows[0];
    for (var k = 1; k < g.rows.length; k++) {
      rowsToDelete[g.rows[k]._row] = true;
    }
    base['Start Date'] = g.start;
    base['End Date'] = g.end;
    base._touchDates = true;
    if (entryIdx >= 0) {
      base['Entry Code'] = String(base['Entry Code'] || '').replace(/-CY$/i, '').replace(/-S[12]$/i, '');
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

  working.sort(function (a, b) {
    var ae = String(a['Emp ID'] || '').toUpperCase();
    var be = String(b['Emp ID'] || '').toUpperCase();
    if (ae !== be) return ae < be ? -1 : 1;
    var as = new Date(a['Start Date']).getTime();
    var bs = new Date(b['Start Date']).getTime();
    if (as !== bs) return as - bs;
    return (a._row || 0) - (b._row || 0);
  });

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

    function applyBalanceYear_(seg) {
      var year = seg.entitlementYear;
      if (!usesAnnualPool) return year;
      if (year >= calYear) return year;
      var rem = getRemainingPrev_(empId, calYear);
      if (rem > 0) return year;
      return calYear;
    }

    if (segments.length === 1) {
      var util = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var days = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      var year1 = applyBalanceYear_(segments[0]);

      if (usesAnnualPool && year1 === calYear - 1) {
        var rem = getRemainingPrev_(empId, calYear);
        if (util > rem && rem > 0) {
          var utilPrev = rem;
          var utilCurr = util - rem;
          setRemainingPrev_(empId, calYear, 0);
          splitCount++;
          var baseCode = String(lv['Entry Code'] || 'LV').replace(/-CY$/i, '').replace(/-S[12]$/i, '');
          updates.push({
            row: lv._row,
            start: segments[0].start,
            end: segments[0].end,
            util: utilPrev,
            days: days,
            year: calYear - 1,
            touchDates: !!lv._touchDates,
            entryCode: baseCode
          });
          var newRow = buildBlankLeaveRow_(headers);
          for (var c = 0; c < headers.length; c++) {
            newRow[c] = lv[headers[c]] !== undefined ? lv[headers[c]] : '';
          }
          if (entryIdx >= 0) newRow[entryIdx] = baseCode + '-CY';
          if (startIdx >= 0) {
            newRow[startIdx] = (typeof toSheetDateValue_ === 'function')
              ? toSheetDateValue_(segments[0].start) : segments[0].start;
          }
          if (endIdx >= 0) {
            newRow[endIdx] = (typeof toSheetDateValue_ === 'function')
              ? toSheetDateValue_(segments[0].end) : segments[0].end;
          }
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
        touchDates: !!lv._touchDates,
        entryCode: String(lv['Entry Code'] || '').replace(/-CY$/i, '').replace(/-S[12]$/i, '') || lv['Entry Code']
      });
    } else {
      splitCount++;
      var u0 = calculateLeaveUtilizeWithMaster_(master, empId, segments[0].start, segments[0].end, lt);
      var d0 = Math.round((segments[0].end - segments[0].start) / 86400000) + 1;
      var baseCode2 = String(lv['Entry Code'] || 'LV').replace(/-CY$/i, '').replace(/-S[12]$/i, '');
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
        touchDates: true,
        entryCode: baseCode2 + '-S1'
      });

      var u1 = calculateLeaveUtilizeWithMaster_(master, empId, segments[1].start, segments[1].end, lt);
      var d1 = Math.round((segments[1].end - segments[1].start) / 86400000) + 1;
      var y1 = segments[1].entitlementYear;
      var newRow2 = buildBlankLeaveRow_(headers);
      for (var c2 = 0; c2 < headers.length; c2++) {
        newRow2[c2] = lv[headers[c2]] !== undefined ? lv[headers[c2]] : '';
      }
      if (entryIdx >= 0) newRow2[entryIdx] = baseCode2 + '-S2';
      if (startIdx >= 0) {
        newRow2[startIdx] = (typeof toSheetDateValue_ === 'function')
          ? toSheetDateValue_(segments[1].start) : segments[1].start;
      }
      if (endIdx >= 0) {
        newRow2[endIdx] = (typeof toSheetDateValue_ === 'function')
          ? toSheetDateValue_(segments[1].end) : segments[1].end;
      }
      if (utilIdx >= 0) newRow2[utilIdx] = u1;
      if (daysIdx >= 0) newRow2[daysIdx] = d1;
      if (yearIdx >= 0) newRow2[yearIdx] = y1;
      appends.push(newRow2);
    }
  });

  var leaveSheet = master.sheet.leave;
  updates.forEach(function (u) {
    if (u.touchDates) {
      if (startIdx >= 0 && u.start) {
        var sv = (typeof toSheetDateValue_ === 'function') ? toSheetDateValue_(u.start) : u.start;
        leaveSheet.getRange(u.row, startIdx + 1).setValue(sv).setNumberFormat('dd-mmm-yyyy');
      }
      if (endIdx >= 0 && u.end) {
        var ev = (typeof toSheetDateValue_ === 'function') ? toSheetDateValue_(u.end) : u.end;
        leaveSheet.getRange(u.row, endIdx + 1).setValue(ev).setNumberFormat('dd-mmm-yyyy');
      }
    }
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
  Logger.log('=== RECALC END === ' + msg);
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
