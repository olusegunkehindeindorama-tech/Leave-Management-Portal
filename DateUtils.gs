/**
 * DATE-ONLY FIX (Nigeria / Africa/Lagos — UTC+1)
 * =================================================
 * Problem: DB/Excel imports stored calendar dates as datetimes. In WAT they
 * often appear as previous day 23:00 (e.g. intended 25-Apr → 24/04/2026 23:00).
 *
 * Goal: always recover the intended CALENDAR date for Start/End leave fields.
 *
 * Strategy:
 *  1. If a Date has a non-zero time in the evening (hour >= 20), add HOURS_NUDGE
 *     so the value crosses into the intended calendar day, then strip time.
 *  2. Otherwise strip to local Y/M/D (no UTC conversion).
 *  3. On write, store local calendar date at 12:00 noon (immune to ±1–2h shifts)
 *     and format as dd-mmm-yyyy.
 *
 * Paste into Apps Script project (merge into DateUtils.gs or as new file).
 * Run repairTblLeaveStartEndDates() once from the editor to fix existing rows.
 */

/** Hours to add when a stored datetime looks like a TZ artifact (WAT = UTC+1). Minimum 2. */
var DATE_TZ_NUDGE_HOURS_ = 2;

/**
 * Convert any sheet value to a pure local calendar Date (time = 00:00:00 local).
 * Handles the 23:00 previous-day pattern from UTC midnight imports.
 */
function toCalendarDate_(val) {
  if (val === null || val === undefined || val === '') return null;

  // Already a Date from Sheets
  if (val instanceof Date && !isNaN(val.getTime())) {
    var d = new Date(val.getTime());
    var h = d.getHours();
    var mi = d.getMinutes();
    var sec = d.getSeconds();

    // Evening times (20:00–23:59) almost always mean "UTC midnight of next day"
    // displayed in Africa/Lagos. Nudge forward so the calendar day is correct.
    if (h >= 20 || (h === 19 && mi >= 0 && DATE_TZ_NUDGE_HOURS_ >= 2)) {
      d = new Date(d.getTime() + DATE_TZ_NUDGE_HOURS_ * 60 * 60 * 1000);
    } else if (h > 0 || mi > 0 || sec > 0) {
      // Any other non-midnight time: keep same local calendar day.
    }

    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  // Strings — prefer explicit date tokens (no timezone)
  var s = String(val).trim();

  // yyyy-MM-dd or yyyy-MM-ddTHH:mm
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }

  // dd-MMM-yyyy / d-MMM-yy
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (mon) {
    var months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    var mi2 = months[mon[2].toLowerCase()];
    if (mi2 !== undefined) {
      var y = Number(mon[3]);
      if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
      return new Date(y, mi2, Number(mon[1]));
    }
  }

  // dd/MM/yyyy or dd-MM-yyyy (Nigeria style)
  var dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  }

  // Fallback: parse then calendar-strip (with nudge if evening)
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return toCalendarDate_(parsed);

  return null;
}

/** Format as yyyy-MM-dd for UI / APIs (never send raw Date for Start/End). */
function formatDateOnly_(val) {
  var d = toCalendarDate_(val);
  if (!d) {
    if (val === null || val === undefined || val === '') return '';
    if (typeof val === 'string') {
      var ms = String(val).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (ms) return ms[1] + '-' + ms[2] + '-' + ms[3];
    }
    return '';
  }
  return d.getFullYear() + '-' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
    ('0' + d.getDate()).slice(-2);
}

/** Parse any input → local calendar Date at 00:00 (for calculations). */
function parseDateOnly_(val) {
  return toCalendarDate_(val);
}

/**
 * Value to WRITE into Start/End cells: local calendar day at 12:00 noon.
 * Noon is immune to ±1–2 hour timezone display shifts.
 */
function toSheetDateValue_(val) {
  var d = toCalendarDate_(val);
  if (!d) return '';
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
}

function formatDateTime_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (!(val instanceof Date) || isNaN(val.getTime())) return String(val);
  try {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  } catch (e) {
    return String(val);
  }
}

/**
 * ONE-TIME REPAIR: fix Start Date + End Date on tblLeave (and optionally tblLeaveDeleted).
 * - Detects evening times (TZ artifacts)
 * - Rewrites as noon local calendar dates
 * - Applies number format dd-mmm-yyyy
 *
 * Run from Apps Script editor: repairTblLeaveStartEndDates()
 * Optional: repairTblLeaveStartEndDates({ nudgeHours: 2, includeDeleted: true })
 */
function repairTblLeaveStartEndDates(opts) {
  opts = opts || {};
  var nudge = (opts.nudgeHours !== undefined && opts.nudgeHours !== null)
    ? Number(opts.nudgeHours) : DATE_TZ_NUDGE_HOURS_;
  if (isNaN(nudge) || nudge < 2) nudge = 2;
  var prevNudge = DATE_TZ_NUDGE_HOURS_;
  DATE_TZ_NUDGE_HOURS_ = isNaN(nudge) ? 2 : Math.max(2, nudge);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ['tblLeave'];
  if (opts.includeDeleted) sheets.push('tblLeaveDeleted');

  var summary = [];
  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) {
      summary.push(name + ': skipped (missing or empty)');
      return;
    }
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h).trim(); });
    var si = headers.indexOf('Start Date');
    var ei = headers.indexOf('End Date');
    if (si < 0 || ei < 0) {
      summary.push(name + ': missing Start Date / End Date columns');
      return;
    }

    var fixed = 0;
    var samples = [];
    for (var r = 1; r < data.length; r++) {
      var rawS = data[r][si];
      var rawE = data[r][ei];
      if (rawS === '' && rawE === '') continue;

      var newS = toSheetDateValue_(rawS);
      var newE = toSheetDateValue_(rawE);

      var changed = false;
      if (newS && (!(rawS instanceof Date) || formatDateOnly_(rawS) !== formatDateOnly_(newS) ||
          (rawS instanceof Date && (rawS.getHours() !== 12 || rawS.getMinutes() !== 0)))) {
        sh.getRange(r + 1, si + 1).setValue(newS).setNumberFormat('dd-mmm-yyyy');
        changed = true;
      } else if (newS) {
        sh.getRange(r + 1, si + 1).setNumberFormat('dd-mmm-yyyy');
      }
      if (newE && (!(rawE instanceof Date) || formatDateOnly_(rawE) !== formatDateOnly_(newE) ||
          (rawE instanceof Date && (rawE.getHours() !== 12 || rawE.getMinutes() !== 0)))) {
        sh.getRange(r + 1, ei + 1).setValue(newE).setNumberFormat('dd-mmm-yyyy');
        changed = true;
      } else if (newE) {
        sh.getRange(r + 1, ei + 1).setNumberFormat('dd-mmm-yyyy');
      }

      if (changed) {
        fixed++;
        if (samples.length < 5) {
          samples.push({
            row: r + 1,
            fromS: String(rawS),
            toS: formatDateOnly_(newS),
            fromE: String(rawE),
            toE: formatDateOnly_(newE)
          });
        }
      }
    }
    summary.push(name + ': repaired ' + fixed + ' row(s). nudgeHours=' + DATE_TZ_NUDGE_HOURS_);
    if (samples.length) summary.push('  samples: ' + JSON.stringify(samples));
  });

  DATE_TZ_NUDGE_HOURS_ = prevNudge;
  SpreadsheetApp.flush();
  var msg = summary.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return { success: true, message: msg };
}

/**
 * Call this whenever IMPORTING a start/end from CSV/Excel/DarwinBox
 * before writing to tblLeave.
 *
 * Example:
 *   cell.setValue(toSheetDateValue_(rawFromCsv)).setNumberFormat('dd-mmm-yyyy');
 */
function importDateOnly_(raw) {
  return toSheetDateValue_(raw);
}
