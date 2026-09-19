/**
 * CacheService layer + shift map (WIDE tblShift date headers may be Date objects).
 */
var CACHE_TTL_SEC = 180;

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
  try { CacheService.getScriptCache().removeAll(['pol', 'emp_map', 'sb', 'filter_opts']); } catch (e) {}
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

/** Normalize a header cell (Date or string) to yyyy-MM-dd */
function shiftHeaderToKey_(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    try {
      return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
      return cell.getFullYear() + '-' +
        ('0' + (cell.getMonth() + 1)).slice(-2) + '-' +
        ('0' + cell.getDate()).slice(-2);
    }
  }
  var ds = String(cell || '').trim();
  if (ds.length >= 10) {
    var m = ds.match(/(\d{4}-\d{2}-\d{2})/);
    if (m) return m[1];
    return ds.substring(0, 10);
  }
  return ds;
}

/**
 * Shift map for one employee: { 'yyyy-MM-dd': code }
 */
function loadShiftMapForEmp_(empId) {
  var key = 'sh_' + String(empId).toUpperCase();
  var hit = cacheGet_(key);
  if (hit) return hit;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tblShift');
  var map = {};
  if (!sh) return map;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return map;

  var target = String(empId).trim().toUpperCase();

  // Header row only first — detect wide vs long
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var h1 = shiftHeaderToKey_(headerRow[1]);
  var isWide = /^\d{4}-\d{2}-\d{2}$/.test(h1) ||
    (String(headerRow[0] || '').toLowerCase().indexOf('emp') === 0 &&
     String(headerRow[1] || '').toLowerCase() !== 'date');

  if (isWide) {
    var dateHeaders = [];
    for (var c = 1; c < headerRow.length; c++) {
      dateHeaders.push(shiftHeaderToKey_(headerRow[c]));
    }
    // Emp ID column — scan rows for match (only col A + full row when found)
    var empCol = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    var foundRow = -1;
    for (var r = 0; r < empCol.length; r++) {
      if (String(empCol[r][0] || '').trim().toUpperCase() === target) {
        foundRow = r + 2;
        break;
      }
    }
    if (foundRow > 0) {
      var rowVals = sh.getRange(foundRow, 1, 1, lastCol).getValues()[0];
      for (var c2 = 1; c2 < rowVals.length; c2++) {
        var code = String(rowVals[c2] || '').trim().toUpperCase();
        if (code && dateHeaders[c2 - 1]) map[dateHeaders[c2 - 1]] = code;
      }
    }
  } else {
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() !== target) continue;
      var d = data[i][1];
      var k = shiftHeaderToKey_(d);
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
    // Keep emp_map — only drop on full sync
  } catch (e) {}
}

/** Unique BU + Department lists for multi-select filters */
function getFilterOptions() {
  var hit = cacheGet_('filter_opts');
  if (hit) return hit;
  var empMap = loadEmployeeMapCached_();
  var bus = {};
  var depts = {};
  Object.keys(empMap).forEach(function (id) {
    if (id === '_headers') return;
    var e = empMap[id];
    var bu = String(e['Business Unit'] || '').trim();
    var dept = String(e['Department'] || '').trim();
    if (bu) bus[bu] = true;
    if (dept) depts[dept] = true;
  });
  var out = {
    bus: Object.keys(bus).sort(),
    departments: Object.keys(depts).sort()
  };
  cachePut_('filter_opts', out);
  return out;
}
