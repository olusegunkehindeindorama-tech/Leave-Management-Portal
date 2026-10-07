/** See artifacts Leave_balance.gs — full content applied via deploy */
function yearChargeLeaveTypes_() {
  return ['Annual Leave', 'Casual Leave', 'Examination Leave', 'Probation Leave'];
}
function yearColNames_(year) {
  return { ent: year + ' Entitlement', util: year + ' Utilized', bal: year + ' Balance' };
}
function computeYearRollup_(typeMeta, usage, year) {
  function utilOf(lt, y) {
    var u = usage[lt] || {};
    return Number(u[y]) || 0;
  }
  var annualEnt = 0, probationEnt = 0;
  if (typeMeta['Annual Leave'] && !typeMeta['Annual Leave'].isUnlimited)
    annualEnt = Number(typeMeta['Annual Leave'].entitlement) || 0;
  if (typeMeta['Probation Leave'] && !typeMeta['Probation Leave'].isUnlimited)
    probationEnt = Number(typeMeta['Probation Leave'].entitlement) || 0;
  var entitlement = annualEnt + probationEnt;
  var utilized = 0;
  var types = yearChargeLeaveTypes_();
  for (var i = 0; i < types.length; i++) utilized += utilOf(types[i], year);
  return { entitlement: entitlement, utilized: utilized, balance: entitlement - utilized };
}
function csvEscape_(val) {
  if (val === null || val === undefined) return '';
  var s = String(val);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
