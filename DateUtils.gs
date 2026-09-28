/**
 * Calendar dates (Start/End) — locale-independent yyyy-MM-dd.
 * Date Entered / Modified / Upload — full datetime in script timezone.
 *
 * CRITICAL: Never send raw Date objects to the browser for Start/End.
 * Always format on the server with formatDateOnly_.
 */

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
        return y + '-' + ('0' + (mi + 1)).slice(-2) + '-' + ('0' + Number(mon[1])).slice(-2);
      }
    }
    var dmy = String(val).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      return dmy[3] + '-' + ('0' + dmy[2]).slice(-2) + '-' + ('0' + dmy[1]).slice(-2);
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

function parseDateOnly_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    var key = formatDateOnly_(val);
    if (key) {
      var p = key.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (p) return new Date(Number(p[1]), Number(p[2]) - 1, Number(p[3]));
    }
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  if (typeof parseLeaveDate_ === 'function') {
    var q = parseLeaveDate_(val);
    if (q) return new Date(q.getFullYear(), q.getMonth(), q.getDate());
  }
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (mon) {
    var months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    var mi = months[mon[2].toLowerCase()];
    if (mi !== undefined) {
      var y = Number(mon[3]);
      if (y < 100) y = y >= 70 ? 1900 + y : 2000 + y;
      return new Date(y, mi, Number(mon[1]));
    }
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
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
