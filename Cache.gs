/**
 * Lightweight CacheService layer.
 * loadShiftMapForEmp_ supports WIDE tblShift (Emp ID | date cols) and legacy LONG.
 */
var CACHE_TTL_SEC = 90;

function cacheGet_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function cachePut_(key, value) {
  try {
    var s = JSON.stringify(value);
    if (s.length < 90000) CacheService.getScriptCache().put(key, s, CACHE_TTL_SEC);
  } catch (e) {}
}

function cacheClearAll_() {
  try { CacheService.getScriptCache().removeAll(['pol', 'emp_all', 'sb']); } catch (e) {}
}

function loadPoliciesCached_() {
  var hit = cacheGet_('pol');
  if (hit) return hit;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Sys_LeavePolicies');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = data[i][c];
    rows.push(o);
  }
  cachePut_('pol', rows);
  return rows;
}

function loadEmployeeMapCached_() {
  var hit = cacheGet_('emp_map');
  if (hit) return hit;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var idIdx = headers.indexOf('Emp ID');
  if (idIdx < 0) idIdx = 1;
  var map = { _headers: headers };
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idIdx]).trim().toUpperCase();
    if (!id) continue;
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = data[i][c];
    map[id] = o;
  }
  cachePut_('emp_map', map);
  return map;
}

/**
 * Shift map for one employee: { 'yyyy-MM-dd': code }
 * WIDE sheet: one row per emp — O(rows) find, then O(dates) fill.
 * LONG sheet (legacy): scan matching emp rows.
 */
function loadShiftMapForEmp_(empId) {
  var key = 'sh_' + String(empId).toUpperCase();
  var hit = cacheGet_(key);
  if (hit) return hit;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblShift');
  var map = {};
  if (!sh) return map;

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return map;

  var target = String(empId).trim().toUpperCase();
  var h1 = String(data[0][1] || '').trim();
  var isWide = /^\d{4}-\d{2}-\d{2}/.test(h1) ||
    (String(data[0][0] || '').toLowerCase().indexOf('emp') === 0 && h1.toLowerCase() !== 'date');

  if (isWide) {
    var dateHeaders = [];
    for (var c = 1; c < data[0].length; c++) {
      var ds = String(data[0][c] || '').trim();
      if (ds.length >= 10) ds = ds.substring(0, 10);
      dateHeaders.push(ds);
    }
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0] || '').trim().toUpperCase() !== target) continue;
      for (var c2 = 1; c2 < data[r].length; c2++) {
        var code = String(data[r][c2] || '').trim().toUpperCase();
        if (code && dateHeaders[c2 - 1]) map[dateHeaders[c2 - 1]] = code;
      }
      break; // one row per emp
    }
  } else {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() !== target) continue;
      var d = data[i][1];
      var k;
      if (d instanceof Date && !isNaN(d.getTime())) {
        k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
      } else {
        k = String(d || '').trim().substring(0, 10);
      }
      if (k) map[k] = String(data[i][2] || '').trim().toUpperCase();
    }
  }

  cachePut_(key, map);
  return map;
}

function loadLeaveRowsForEmp_(empId) {
  var key = 'lv_' + String(empId).toUpperCase();
  var hit = cacheGet_(key);
  if (hit) return hit;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblLeave');
  var out = [];
  if (!sh) return out;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return out;
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var empIdx = headers.indexOf('Emp ID');
  var target = String(empId).trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][empIdx]).trim().toUpperCase() !== target) continue;
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = data[i][c];
    out.push(o);
  }
  cachePut_(key, out);
  return out;
}

function invalidateEmpCaches_(empId) {
  try {
    var c = CacheService.getScriptCache();
    c.remove('lv_' + String(empId).toUpperCase());
    c.remove('sh_' + String(empId).toUpperCase());
    c.remove('emp_map');
  } catch (e) {}
}
