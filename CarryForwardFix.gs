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
