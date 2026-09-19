/**
 * ============================================================
 *  LEAVE RECORD CLEANUP (SAFE — in-place)
 * ============================================================
 *  Dedup key (strict, leave-code ignored):
 *      EmpID | yyyy-MM-dd | yyyy-MM-dd
 *  Dates compared via script timezone so 09-Jun-2026 as text/Date/serial
 *  all map to the same key.
 *
 *  NEVER clears the sheet. NEVER drops rows that cannot form a key.
 *  Deletes only true duplicate sheet rows (bottom-up).
 *
 *  Run:
 *    cleanupDuplicateLeaveRecordsOnly()  — dedupe only
 *    cleanupDuplicateLeaveRecords()      — dedupe then calculateLeaveUtilized
 * ============================================================
 */

/** Timezone-safe yyyy-MM-dd (script TZ). */
function leaveDateKey_(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    var y = d.getFullYear();
    var m = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);
    return y + '-' + m + '-' + day;
  }
}

/** Parse any cell value → Date at local calendar day, or null. */
function leaveParseDate_(val) {
  if (typeof parseLeaveDate_ === 'function') {
    return parseLeaveDate_(val);
  }
  // Minimal fallback if DB Import not loaded
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).trim();
  var months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (mon && months[mon[2].toLowerCase()] !== undefined) {
    var y = Number(mon[3]); if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
    return new Date(y, months[mon[2].toLowerCase()], Number(mon[1]));
  }
  var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slash) {
    var yy = Number(slash[3]); if (yy < 100) yy = yy >= 70 ? 1900 + yy : 2000 + yy;
    return new Date(yy, Number(slash[2]) - 1, Number(slash[1]));
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function leaveFingerprint_(empId, startDate, endDate) {
  return String(empId).trim().toUpperCase() + '|' +
    leaveDateKey_(startDate) + '|' + leaveDateKey_(endDate);
}

function stripEntryCodeSuffix_(code) {
  var c = String(code || '').trim();
  if (!c) return c;
  var prev;
  do {
    prev = c;
    c = c.replace(/-S[12]$/i, '');
    c = c.replace(/-[ab]$/i, '');
  } while (c !== prev);
  return c;
}

function leaveRowQualityScore_(row, headers) {
  var score = 0;
  function col(name) {
    var i = headers.indexOf(name);
    return i >= 0 ? row[i] : '';
  }
  var entry = String(col('Entry Code') || '');
  var leaveCode = String(col('Leave Code') || '').trim();
  var name = String(col('Emp Name') || '').trim();
  var util = col('Leave Utilized');
  var days = col('No of Days');
  var reason = String(col('Leave Reason') || '').trim();
  var enteredBy = String(col('Entered By') || '').trim();

  if (leaveCode && leaveCode.toUpperCase() !== 'NA') score += 10;
  if (name) score += 5;
  if (util !== '' && util !== null && util !== undefined) score += 3;
  if (days !== '' && days !== null && days !== undefined) score += 2;
  if (reason) score += 1;
  if (entry && entry === stripEntryCodeSuffix_(entry)) score += 4;
  // Prefer BP- human codes over auto DB- codes when both exist for same dates
  if (/^BP-/i.test(entry)) score += 6;
  if (/^DB-\d+/i.test(entry)) score -= 2;
  if (enteredBy && enteredBy.toLowerCase() !== 'darwinbox') score += 1;
  return score;
}

/**
 * @param {boolean} optSkipRecalc  true = dedupe only
 */
function cleanupDuplicateLeaveRecords(optSkipRecalc) {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    return { success: false, message: 'tblLeave sheet missing.' };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return { success: true, message: 'No leave rows to clean.', removed: 0 };
  }

  // getValues + getDisplayValues so we can parse either real dates or text
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');

  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'tblLeave missing Emp ID / Start Date / End Date.' };
  }

  // Build groups: leave-code is intentionally ignored
  var groups = {}; // fp → [{ sheetRow, dataIdx, score, startNorm, endNorm }]
  var ungrouped = 0;
  var sampleKeys = [];

  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdx] || '').trim().toUpperCase();
    // Prefer parsing raw value; fall back to display text (handles text-formatted dates)
    var s = leaveParseDate_(data[i][startIdx]) || leaveParseDate_(display[i][startIdx]);
    var e = leaveParseDate_(data[i][endIdx]) || leaveParseDate_(display[i][endIdx]);

    if (!empId || !s || !e) {
      ungrouped++;
      continue;
    }

    var fp = leaveFingerprint_(empId, s, e);
    if (!groups[fp]) groups[fp] = [];
    groups[fp].push({
      sheetRow: i + 1,
      dataIdx: i,
      score: leaveRowQualityScore_(data[i], headers),
      startNorm: s,
      endNorm: e
    });
    if (sampleKeys.length < 5 && groups[fp].length > 1) {
      sampleKeys.push(fp + ' x' + groups[fp].length);
    }
  }

  var rowsToDelete = [];
  var rowsToUpdate = [];
  var dupGroups = 0;

  Object.keys(groups).forEach(function (fp) {
    var list = groups[fp];
    list.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.sheetRow - b.sheetRow;
    });

    var keep = list[0];
    var upd = { sheetRow: keep.sheetRow };
    var changed = false;

    if (entryIdx >= 0) {
      var raw = String(data[keep.dataIdx][entryIdx] || '');
      var clean = stripEntryCodeSuffix_(raw);
      if (clean !== raw) {
        upd.entryCode = clean;
        changed = true;
      }
    }
    // Always write normalized Date so format is consistent in the sheet
    upd.start = keep.startNorm;
    upd.end = keep.endNorm;
    changed = true;
    if (changed) rowsToUpdate.push(upd);

    if (list.length > 1) {
      dupGroups++;
      for (var d = 1; d < list.length; d++) {
        rowsToDelete.push(list[d].sheetRow);
      }
    }
  });

  // Apply updates before deletes (row numbers still original)
  rowsToUpdate.forEach(function (u) {
    if (u.entryCode !== undefined && entryIdx >= 0) {
      sheet.getRange(u.sheetRow, entryIdx + 1).setValue(u.entryCode);
    }
    if (u.start !== undefined) {
      sheet.getRange(u.sheetRow, startIdx + 1).setValue(u.start);
    }
    if (u.end !== undefined) {
      sheet.getRange(u.sheetRow, endIdx + 1).setValue(u.end);
    }
  });

  // Delete duplicates bottom-up in contiguous blocks
  rowsToDelete.sort(function (a, b) { return b - a; });
  var deleted = 0;
  var di = 0;
  while (di < rowsToDelete.length) {
    var blockEnd = rowsToDelete[di];
    var blockStart = blockEnd;
    while (di + 1 < rowsToDelete.length && rowsToDelete[di + 1] === blockStart - 1) {
      di++;
      blockStart = rowsToDelete[di];
    }
    sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
    deleted += (blockEnd - blockStart + 1);
    di++;
  }

  SpreadsheetApp.flush();

  // Uniform display format for date columns
  try {
    var newLast = sheet.getLastRow();
    if (newLast >= 2) {
      sheet.getRange(2, startIdx + 1, newLast - 1, 1).setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(2, endIdx + 1, newLast - 1, 1).setNumberFormat('dd-mmm-yyyy');
    }
  } catch (fmtErr) {}

  if (typeof cacheClearAll_ === 'function') cacheClearAll_();

  var recalcResult = null;
  if (!optSkipRecalc && typeof calculateLeaveUtilized === 'function') {
    try {
      recalcResult = calculateLeaveUtilized();
    } catch (re) {
      recalcResult = { success: false, message: 'Recalc error: ' + re.message };
    }
  }

  var ms = new Date().getTime() - started;
  var msg = 'Cleanup in ' + ms + ' ms: ' + dupGroups + ' duplicate group(s), removed ' +
    deleted + ' row(s), normalized ' + rowsToUpdate.length + ' kept row(s). ' +
    ungrouped + ' row(s) skipped (no Emp ID or unparseable dates — kept).';
  if (sampleKeys.length) msg += ' Sample dups: ' + sampleKeys.join('; ') + '.';
  if (recalcResult && recalcResult.message) msg += ' | Recalc: ' + recalcResult.message;

  Logger.log(msg);
  return {
    success: true,
    message: msg,
    removed: deleted,
    dupGroups: dupGroups,
    updated: rowsToUpdate.length,
    ungroupedKept: ungrouped,
    sampleDupKeys: sampleKeys,
    recalc: recalcResult,
    elapsedMs: ms
  };
}

function cleanupLeaveDuplicates() {
  return cleanupDuplicateLeaveRecords(false);
}

/** Dedupe only — no split/recalc. Run this first and verify FRT7351. */
function cleanupDuplicateLeaveRecordsOnly() {
  return cleanupDuplicateLeaveRecords(true);
}
