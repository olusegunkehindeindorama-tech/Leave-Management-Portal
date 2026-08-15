/**
 * Lightweight in-memory / CacheService layer to avoid repeated full-sheet reads.
 * Cache TTL: 90 seconds (short enough for multi-user freshness).
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

/** Policy rows as plain objects (small). */
function loadPoliciesCached_() {
  var hit = cacheGet_('pol');
  if (hit) return hit;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Sys_LeavePolicies');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = data[i][c];
    rows.push(o);
  }
  cachePut_('pol', rows);
  return rows;
}

/** Single employee row by ID (scans emp sheet once, caches map of id->row). */
function loadEmployeeMapCached_() {
  var hit = cacheGet_('emp_map');
  if (hit) return hit;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblEmployee') || ss.getSheetByName('tblemployee');
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
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

/** Shift map for one employee: { 'yyyy-MM-dd': code } — only that emp's rows. */
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
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() !== target) continue;
    var d = new Date(data[i][1]);
    if (isNaN(d.getTime())) continue;
    var k = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    map[k] = String(data[i][2] || '').trim().toUpperCase();
  }
  cachePut_(key, map);
  return map;
}

/** Leave rows for one emp only (filter in memory after one read; cache per emp). */
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
  var headers = data[0].map(function(h) { return String(h).trim(); });
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
