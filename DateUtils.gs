/**
 * Calendar dates (Start/End) vs datetimes (Entered/Modified/Upload).
 *
 * Start Date / End Date:
 *   - Stored and displayed as pure calendar dates (yyyy-MM-dd)
 *   - Never use toISOString / UTC — that shifts the day in Nigeria (WAT)
 *
 * Date Entered / Date Modified / Upload Date:
 *   - Full datetime in script timezone
 */

/** Format any sheet Date or string as yyyy-MM-dd using script timezone (not UTC). */
function formatDateOnly_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'string') {
    var ms = String(val).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (ms) return ms[1] + '-' + ms[2] + '-' + ms[3];
    var mon = String(val).trim().match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
    if (mon) {
      var months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
      var mi = months[mon[2].toLowerCase()];
      if (mi !== undefined) {
        var y = Number(mon[3]);
        if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
        var dtmp = new Date(y, mi, Number(mon[1]));
        return Utilities.formatDate(dtmp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    }
  }
  if (val instanceof Date && !isNaN(val.getTime())) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
      return val.getFullYear() + '-' +
        ('0' + (val.getMonth() + 1)).slice(-2) + '-' +
        ('0' + val.getDate()).slice(-2);
    }
  }
  return '';
}

/**
 * Parse UI/sheet value into a Date at local midnight (year, month, day only).
 * Avoids `new Date('yyyy-MM-dd')` which is UTC midnight and shifts in WAT.
 */
function parseDateOnly_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  if (typeof parseLeaveDate_ === 'function') {
    var p = parseLeaveDate_(val);
    if (p) return new Date(p.getFullYear(), p.getMonth(), p.getDate());
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

/** Format datetime in script timezone (for Date Entered / Modified / Upload). */
function formatDateTime_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (!(val instanceof Date) || isNaN(val.getTime())) return String(val);
  try {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  } catch (e) {
    return String(val);
  }
}
