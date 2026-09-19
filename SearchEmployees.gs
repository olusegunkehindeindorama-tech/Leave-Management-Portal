/**
 * Search employees by Emp ID OR Emp Name (partial match).
 * NOTE: If Code.gs also defines searchEmployees, remove the Code.gs copy
 * so this definition is used (Apps Script does not allow duplicates).
 */
function searchEmployees(query, limit) {
  var empMap = loadEmployeeMapCached_();
  var needle = String(query || '').trim().toLowerCase();
  var max = Number(limit) || 12;
  if (!needle) return [];
  var matches = [];
  var ids = Object.keys(empMap);
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    if (id === '_headers') continue;
    var e = empMap[id];
    var name = String(e['Emp Name'] || '').toLowerCase();
    if (id.toLowerCase().indexOf(needle) !== -1 || name.indexOf(needle) !== -1) {
      matches.push({
        id: id,
        name: String(e['Emp Name'] || ''),
        department: String(e['Department'] || ''),
        bu: String(e['Business Unit'] || '')
      });
      if (matches.length >= max * 3) break; // collect extra then sort/trim
    }
  }
  matches.sort(function (a, b) {
    var ae = a.id.toLowerCase() === needle ? 0 : (a.id.toLowerCase().indexOf(needle) === 0 ? 1 : 2);
    var be = b.id.toLowerCase() === needle ? 0 : (b.id.toLowerCase().indexOf(needle) === 0 ? 1 : 2);
    if (ae !== be) return ae - be;
    return a.name.localeCompare(b.name);
  });
  return matches.slice(0, max);
}

/** Multi-value filters: "A|B|C" or single string */
function matchFilterMulti_(cellValue, filterStr) {
  if (!filterStr) return true;
  var cell = String(cellValue || '').trim().toUpperCase();
  var parts = String(filterStr).split('|').map(function (s) {
    return s.trim().toUpperCase();
  }).filter(Boolean);
  if (!parts.length) return true;
  return parts.indexOf(cell) !== -1;
}
