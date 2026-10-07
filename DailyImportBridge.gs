/**
 * DailyImportBridge.gs — after DailyLeaveModel.gs
 * range leave objects → daily rows.
 */

function importRangeLeavesAsDaily_(rangeLeaves, sourceLabel) {
  Logger.log('importRangeLeavesAsDaily_ from ' + (sourceLabel || 'source') +
    ': ' + (rangeLeaves ? rangeLeaves.length : 0) + ' range apps');
  ensureDailyTblLeaveSchema_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  var allDaily = [];
  for (var i = 0; i < (rangeLeaves || []).length; i++) {
    var L = rangeLeaves[i];
    var rows = expandLeaveApplicationToDailyRows_(
      L.entryCode || L.entry || '',
      L.empId || L.emp || '',
      L.leaveType || L.type || '',
      L.startDate || L.start,
      L.endDate || L.end || L.startDate || L.start,
      L.leaveReason || L.reason || ''
    );
    for (var j = 0; j < rows.length; j++) allDaily.push(rows[j]);
  }
  var result = appendDailyLeaveRows_(sheet, allDaily, null);
  Logger.log('importRangeLeavesAsDaily_ appended=' + result.appended +
    ' skippedDup=' + result.skipped);
  return {
    success: true,
    message: (sourceLabel || 'Import') + ': +' + result.appended +
      ' daily rows (skipped dup ' + result.skipped + ')',
    appended: result.appended,
    skipped: result.skipped
  };
}

function submitLeaveApplicationDaily(payload) {
  if (!payload || !payload.empId || !payload.startDate) {
    return { success: false, message: 'empId and startDate required' };
  }
  var entry = payload.entryCode || (
    'UI-' + String(payload.empId).toUpperCase() + '-' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')
  );
  var rows = expandLeaveApplicationToDailyRows_(
    entry,
    payload.empId,
    payload.leaveType || 'Annual Leave',
    payload.startDate,
    payload.endDate || payload.startDate,
    payload.leaveReason || ''
  );
  ensureDailyTblLeaveSchema_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tblLeave');
  var result = appendDailyLeaveRows_(sheet, rows, null);
  if (result.appended && typeof calculateLeaveUtilized === 'function') {
    try { calculateLeaveUtilized(); } catch (e) {
      Logger.log('post-submit recalc: ' + e.message);
    }
  }
  return {
    success: result.appended > 0,
    message: result.appended
      ? ('Leave saved: ' + result.appended + ' day(s)' +
          (result.skipped ? ', ' + result.skipped + ' day(s) already existed' : ''))
      : ('No new days saved (all ' + result.skipped + ' already on file)'),
    entryCode: entry,
    appended: result.appended,
    skipped: result.skipped
  };
}
