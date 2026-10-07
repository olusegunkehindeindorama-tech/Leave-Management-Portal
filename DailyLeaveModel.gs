/**
 * ============================================================
 *  DAILY LEAVE MODEL (canonical tblLeave schema)
 * ============================================================
 *  Columns:
 *    Entry Code | Emp ID | Leave Type | Leave Date |
 *    Leave Utilized | Entitlement Year | Leave Reason
 *
 *  One row = one employee on leave for one calendar day.
 *  Multi-day applications expand Start→End into N daily rows
 *  sharing the same Entry Code.
 *
 *  Dedup key: EmpID | yyyy-MM-dd | Leave Type
 *  Leave Utilized = shift multiplier for that day (0 / 1 / 1.5 …)
 * ============================================================
 */

var DAILY_LEAVE_HEADERS = [
  'Entry Code',
  'Emp ID',
  'Leave Type',
  'Leave Date',
  'Leave Utilized',
  'Entitlement Year',
  'Leave Reason'
];

function ensureDailyTblLeaveSchema_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) sheet = ss.insertSheet('tblLeave');
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var needRewrite = DAILY_LEAVE_HEADERS.some(function (h, i) {
    return existing[i] !== h;
  }) || existing.length < DAILY_LEAVE_HEADERS.length;

  if (needRewrite && sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, DAILY_LEAVE_HEADERS.length).setValues([DAILY_LEAVE_HEADERS]);
    sheet.setFrozenRows(1);
    Logger.log('ensureDailyTblLeaveSchema_: headers written (empty sheet)');
  } else if (needRewrite) {
    Logger.log('ensureDailyTblLeaveSchema_: sheet has data with old headers — run migrateTblLeaveToDaily()');
  }
  return sheet;
}

function isDailyTblLeaveSchema_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return false;
  var h = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 12)).getValues()[0]
    .map(function (x) { return String(x || '').trim(); });
  return h.indexOf('Leave Date') >= 0 && h.indexOf('Start Date') < 0;
}

function dailyDateKey_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    var d = val;
    if (d.getHours() >= 20) d = new Date(d.getTime() + 2 * 60 * 60 * 1000);
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2);
  }
  var s = String(val).trim();
  if (s.charAt(0) === "'") s = s.substring(1);
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  var months = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  if (mon && months[mon[2].toLowerCase()]) {
    var y = Number(mon[3]); if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
    return y + '-' + ('0' + months[mon[2].toLowerCase()]).slice(-2) + '-' +
      ('0' + Number(mon[1])).slice(-2);
  }
  var dmy = s.match(/^(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{2,4})$/);
  if (dmy) {
    var day = Number(dmy[1]), month = Number(dmy[2]), year = Number(dmy[3]);
    if (year < 100) year = year >= 70 ? 1900 + year : 2000 + year;
    if (month > 12 && day <= 12) { var t = day; day = month; month = t; }
    if (month < 1 || month > 12 || day < 1 || day > 31) return '';
    return year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
  }
  return '';
}

function dailySheetDate_(val) {
  var key = dailyDateKey_(val);
  if (!key) return '';
  var p = key.split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}

function expandDateRangeKeys_(startVal, endVal) {
  var sk = dailyDateKey_(startVal);
  var ek = dailyDateKey_(endVal || startVal);
  if (!sk) return [];
  if (!ek) ek = sk;
  if (ek < sk) { var tmp = sk; sk = ek; ek = tmp; }
  var out = [];
  var p = sk.split('-');
  var cur = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
  var endP = ek.split('-');
  var end = new Date(Number(endP[0]), Number(endP[1]) - 1, Number(endP[2]), 12, 0, 0);
  var guard = 0;
  while (cur.getTime() <= end.getTime() && guard < 370) {
    out.push(
      cur.getFullYear() + '-' +
      ('0' + (cur.getMonth() + 1)).slice(-2) + '-' +
      ('0' + cur.getDate()).slice(-2)
    );
    cur.setDate(cur.getDate() + 1);
    guard++;
  }
  return out;
}

function dailyFingerprint_(empId, leaveDateKey, leaveType) {
  return String(empId || '').trim().toUpperCase() + '|' +
    String(leaveDateKey || '') + '|' +
    String(leaveType || '').trim().toUpperCase();
}

function buildDailyLeaveRow_(entryCode, empId, leaveType, leaveDateKey, utilized, entitlementYear, leaveReason) {
  return [
    String(entryCode || ''),
    String(empId || '').trim().toUpperCase(),
    String(leaveType || '').trim(),
    dailySheetDate_(leaveDateKey),
    utilized === '' || utilized === null || utilized === undefined ? '' : Number(utilized),
    entitlementYear === '' || entitlementYear === null || entitlementYear === undefined ? '' : Number(entitlementYear),
    String(leaveReason || '')
  ];
}

function expandLeaveApplicationToDailyRows_(entryCode, empId, leaveType, startVal, endVal, leaveReason) {
  var keys = expandDateRangeKeys_(startVal, endVal);
  var rows = [];
  for (var i = 0; i < keys.length; i++) {
    rows.push(buildDailyLeaveRow_(entryCode, empId, leaveType, keys[i], '', '', leaveReason));
  }
  return rows;
}

function appendDailyLeaveRows_(sheet, dailyRows, existingFpMap) {
  if (!dailyRows || !dailyRows.length) return { appended: 0, skipped: 0 };
  sheet = sheet || ensureDailyTblLeaveSchema_();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var empI = headers.indexOf('Emp ID');
  var dateI = headers.indexOf('Leave Date');
  var typeI = headers.indexOf('Leave Type');

  var fpMap = existingFpMap || {};
  if (!existingFpMap && sheet.getLastRow() >= 2 && empI >= 0 && dateI >= 0 && typeI >= 0) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var r = 0; r < data.length; r++) {
      var ek = dailyDateKey_(data[r][dateI]);
      var fp = dailyFingerprint_(data[r][empI], ek, data[r][typeI]);
      if (fp) fpMap[fp] = true;
    }
  }

  var toWrite = [];
  var skipped = 0;
  for (var i = 0; i < dailyRows.length; i++) {
    var row = dailyRows[i];
    var emp = row[1];
    var dk = dailyDateKey_(row[3]);
    var lt = row[2];
    var fp2 = dailyFingerprint_(emp, dk, lt);
    if (fpMap[fp2]) { skipped++; continue; }
    fpMap[fp2] = true;
    toWrite.push(row);
  }
  if (toWrite.length) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toWrite.length, DAILY_LEAVE_HEADERS.length).setValues(toWrite);
    sheet.getRange(startRow, 4, toWrite.length, 1).setNumberFormat('dd-mmm-yyyy');
    SpreadsheetApp.flush();
  }
  return { appended: toWrite.length, skipped: skipped, fpMap: fpMap };
}

function migrateTblLeaveToDaily() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) return { success: false, message: 'tblLeave missing' };

  if (isDailyTblLeaveSchema_(sheet)) {
    return { success: true, message: 'Already daily schema — nothing to migrate', rows: sheet.getLastRow() - 1 };
  }

  Logger.log('=== migrateTblLeaveToDaily START ===');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    ensureDailyTblLeaveSchema_();
    return { success: true, message: 'Empty sheet — headers set to daily', rows: 0 };
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h || '').trim(); });
  var iEntry = headers.indexOf('Entry Code');
  var iEmp = headers.indexOf('Emp ID');
  var iType = headers.indexOf('Leave Type');
  var iStart = headers.indexOf('Start Date');
  var iEnd = headers.indexOf('End Date');
  var iReason = headers.indexOf('Leave Reason');
  var iUtil = headers.indexOf('Leave Utilized');
  var iYear = headers.indexOf('Entitlement Year');

  if (iEmp < 0 || iStart < 0) {
    return { success: false, message: 'Legacy columns Emp ID / Start Date missing' };
  }

  var expanded = [];
  var fpMap = {};
  var sourceRows = 0;
  var dupSkip = 0;

  for (var r = 1; r < data.length; r++) {
    var emp = String(data[r][iEmp] || '').trim().toUpperCase();
    if (!emp) continue;
    sourceRows++;
    var entry = iEntry >= 0 ? String(data[r][iEntry] || '') : '';
    var lt = iType >= 0 ? String(data[r][iType] || '') : '';
    var reason = iReason >= 0 ? String(data[r][iReason] || '') : '';
    var start = data[r][iStart];
    var end = iEnd >= 0 ? data[r][iEnd] : start;
    var keys = expandDateRangeKeys_(start, end);
    if (!keys.length) continue;

    var totalUtil = iUtil >= 0 ? Number(data[r][iUtil]) || 0 : 0;
    var year = iYear >= 0 ? data[r][iYear] : '';
    var perDay = keys.length ? totalUtil / keys.length : 0;

    for (var k = 0; k < keys.length; k++) {
      var fp = dailyFingerprint_(emp, keys[k], lt);
      if (fpMap[fp]) { dupSkip++; continue; }
      fpMap[fp] = true;
      expanded.push(buildDailyLeaveRow_(
        entry, emp, lt, keys[k],
        totalUtil ? Math.round(perDay * 1000) / 1000 : '',
        year, reason
      ));
    }
  }

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  var legacyName = 'tblLeave_legacy_' + stamp;
  try {
    sheet.setName(legacyName);
    Logger.log('Renamed old sheet → ' + legacyName);
  } catch (e) {
    Logger.log('Rename failed: ' + e.message);
  }

  var fresh = ss.getSheetByName('tblLeave');
  if (!fresh) fresh = ss.insertSheet('tblLeave');
  fresh.clear();
  fresh.getRange(1, 1, 1, DAILY_LEAVE_HEADERS.length).setValues([DAILY_LEAVE_HEADERS]);
  fresh.setFrozenRows(1);

  if (expanded.length) {
    var chunk = 500;
    for (var i = 0; i < expanded.length; i += chunk) {
      var part = expanded.slice(i, i + chunk);
      fresh.getRange(i + 2, 1, part.length, DAILY_LEAVE_HEADERS.length).setValues(part);
      SpreadsheetApp.flush();
    }
    fresh.getRange(2, 4, expanded.length, 1).setNumberFormat('dd-mmm-yyyy');
  }

  var msg = 'Migrated ' + sourceRows + ' range rows → ' + expanded.length +
    ' daily rows (skipped ' + dupSkip + ' day-dups). Legacy: ' + legacyName;
  Logger.log('=== migrateTblLeaveToDaily END === ' + msg);
  return {
    success: true,
    message: msg,
    sourceRows: sourceRows,
    dailyRows: expanded.length,
    skippedDayDups: dupSkip,
    legacySheet: legacyName
  };
}
