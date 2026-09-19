/**
 * ============================================================
 *  LEAVE RECORD CLEANUP (on-click)
 * ============================================================
 *  1. Normalize Start/End dates to pure calendar dates
 *  2. Deduplicate by Emp ID + Start Date + End Date (keep one row)
 *  3. Strip messy Entry Code suffixes (-a, -b, -S1, -S2, -a-S1, …)
 *     to a clean base code; later recalc will re-apply -S1/-S2 only
 *  4. Rewrite tblLeave uniquely, then run calculateLeaveUtilized
 *
 *  Run: cleanupDuplicateLeaveRecords()
 * ============================================================
 */

/**
 * Strip legacy / split suffixes from an Entry Code.
 * Examples:
 *   BP-1023-a     → BP-1023
 *   BP-1023-b     → BP-1023
 *   BP-11073-S1   → BP-11073
 *   BP-11287-a-S1 → BP-11287
 *   BP-11287-b-S2 → BP-11287
 */
function stripEntryCodeSuffix_(code) {
  var c = String(code || '').trim();
  if (!c) return c;
  // Repeatedly strip trailing -S1/-S2/-a/-b (case-insensitive)
  var prev;
  do {
    prev = c;
    c = c.replace(/-S[12]$/i, '');
    c = c.replace(/-[ab]$/i, '');
  } while (c !== prev);
  return c;
}

/**
 * Preference score for which duplicate row to keep (higher = better).
 */
function leaveRowQualityScore_(row, headers) {
  var score = 0;
  var entry = String(row[headers.indexOf('Entry Code')] || '');
  var leaveCode = String(row[headers.indexOf('Leave Code')] || '').trim();
  var name = String(row[headers.indexOf('Emp Name')] || '').trim();
  var util = row[headers.indexOf('Leave Utilized')];
  var days = row[headers.indexOf('No of Days')];
  var reason = String(row[headers.indexOf('Leave Reason')] || '').trim();
  var enteredBy = String(row[headers.indexOf('Entered By')] || '').trim();

  if (leaveCode && leaveCode.toUpperCase() !== 'NA') score += 10;
  if (name) score += 5;
  if (util !== '' && util !== null && util !== undefined) score += 3;
  if (days !== '' && days !== null && days !== undefined) score += 2;
  if (reason) score += 1;
  // Prefer simpler entry codes (no suffix)
  if (entry === stripEntryCodeSuffix_(entry)) score += 4;
  // Prefer non-random DB auto codes slightly less than human codes
  if (/^DB-\d+/i.test(entry)) score -= 1;
  if (enteredBy && enteredBy.toLowerCase() !== 'darwinbox') score += 1;
  return score;
}

/**
 * Main cleanup entry point.
 * @param {boolean} optSkipRecalc  if true, only dedupe/normalize (no split/recalc)
 */
function cleanupDuplicateLeaveRecords(optSkipRecalc) {
  var started = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('tblLeave');
  if (!sheet) {
    return { success: false, message: 'tblLeave sheet missing.' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { success: true, message: 'No leave rows to clean.', kept: 0, removed: 0 };
  }

  var headers = data[0].map(function (h) { return String(h).trim(); });
  var empIdx = headers.indexOf('Emp ID');
  var startIdx = headers.indexOf('Start Date');
  var endIdx = headers.indexOf('End Date');
  var entryIdx = headers.indexOf('Entry Code');

  if (empIdx < 0 || startIdx < 0 || endIdx < 0) {
    return { success: false, message: 'tblLeave missing Emp ID / Start Date / End Date.' };
  }

  // Group by fingerprint
  var groups = {}; // fp → [ { rowIndex, values, score } ]
  var badRows = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i].slice();
    var empId = String(row[empIdx] || '').trim().toUpperCase();
    var s = parseLeaveDate_(row[startIdx]);
    var e = parseLeaveDate_(row[endIdx]);

    if (!empId || !s || !e) {
      badRows++;
      continue; // drop unparseable
    }

    // Normalize dates in the row values
    row[startIdx] = s;
    row[endIdx] = e;
    row[empIdx] = empId;

    // Clean entry code base (suffixes stripped; -S1/-S2 re-applied later by recalc)
    if (entryIdx >= 0) {
      row[entryIdx] = stripEntryCodeSuffix_(row[entryIdx]);
    }

    var fp = fingerprintKey_(empId, s, e);
    if (!groups[fp]) groups[fp] = [];
    groups[fp].push({
      rowIndex: i + 1,
      values: row,
      score: leaveRowQualityScore_(row, headers)
    });
  }

  var kept = [];
  var removed = 0;
  var fps = Object.keys(groups);

  for (var g = 0; g < fps.length; g++) {
    var list = groups[fps[g]];
    // Sort by score desc, then by rowIndex asc (stable preference)
    list.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.rowIndex - b.rowIndex;
    });
    kept.push(list[0].values);
    removed += list.length - 1;
  }

  // Sort kept rows for readability: Emp ID, Start Date
  kept.sort(function (a, b) {
    var ae = String(a[empIdx] || '');
    var be = String(b[empIdx] || '');
    if (ae !== be) return ae < be ? -1 : 1;
    return new Date(a[startIdx]) - new Date(b[startIdx]);
  });

  // Rewrite sheet: header + unique rows (single write)
  var out = [headers].concat(kept);
  sheet.clearContents();
  // Ensure capacity
  var need = out.length - sheet.getMaxRows();
  if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need + 10);

  var CHUNK = 5000;
  for (var c = 0; c < out.length; c += CHUNK) {
    var part = out.slice(c, c + CHUNK);
    sheet.getRange(c + 1, 1, part.length, headers.length).setValues(part);
  }
  SpreadsheetApp.flush();

  // Format Start/End columns as dates for consistent display
  try {
    if (kept.length > 0) {
      sheet.getRange(2, startIdx + 1, kept.length, 1).setNumberFormat('dd-mmm-yyyy');
      sheet.getRange(2, endIdx + 1, kept.length, 1).setNumberFormat('dd-mmm-yyyy');
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
  var msg = 'Cleanup done in ' + ms + ' ms. Kept ' + kept.length +
    ' unique leave(s), removed ' + removed + ' duplicate(s)' +
    (badRows ? ', dropped ' + badRows + ' unparseable row(s)' : '') +
    '. Entry codes stripped to base (recalc applies -S1/-S2 only).';
  if (recalcResult && recalcResult.message) {
    msg += ' | Recalc: ' + recalcResult.message;
  }
  Logger.log(msg);
  return {
    success: true,
    message: msg,
    kept: kept.length,
    removed: removed,
    badRows: badRows,
    recalc: recalcResult,
    elapsedMs: ms
  };
}

/** Alias */
function cleanupLeaveDuplicates() {
  return cleanupDuplicateLeaveRecords(false);
}
