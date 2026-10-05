/**
 * ImportDateBridge.gs
 * ------------------
 * Load this AFTER: DateUtils.gs, DB Import.gs, Excel Leave Import.gs
 *
 * Hooks Start/End date handling so imports use DateUtils:
 *   - toCalendarDate_ (evening times get min +2h nudge for WAT)
 *   - toSheetDateValue_ (writes local noon → no day shift in sheet)
 *
 * Also hardens parseLeaveDate_ / formatDateKey if DateUtils is present.
 */

(function () {
  if (typeof parseLeaveDate_ === 'function') {
    var _parseLeaveDateOrig = parseLeaveDate_;
    parseLeaveDate_ = function (val) {
      if (typeof toCalendarDate_ === 'function') {
        return toCalendarDate_(val);
      }
      if (val instanceof Date && !isNaN(val.getTime())) {
        var d0 = new Date(val.getTime());
        if (d0.getHours() >= 20) {
          d0 = new Date(d0.getTime() + 2 * 60 * 60 * 1000);
        }
        return new Date(d0.getFullYear(), d0.getMonth(), d0.getDate());
      }
      return _parseLeaveDateOrig(val);
    };
  }

  if (typeof setLeaveCol_ === 'function') {
    var _setLeaveColOrig = setLeaveCol_;
    setLeaveCol_ = function (row, headers, name, value) {
      if (name === 'Start Date' || name === 'End Date') {
        if (value !== null && value !== undefined && value !== '') {
          if (typeof toSheetDateValue_ === 'function') {
            var v = toSheetDateValue_(value);
            value = (v === '' || v === null) ? value : v;
          } else if (value instanceof Date && !isNaN(value.getTime())) {
            var d = value;
            if (d.getHours() >= 20) {
              d = new Date(d.getTime() + 2 * 60 * 60 * 1000);
            }
            value = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
          }
        }
      }
      return _setLeaveColOrig(row, headers, name, value);
    };
  }

  if (typeof formatDateKey === 'function' && typeof formatDateOnly_ === 'function') {
    var _formatDateKeyOrig = formatDateKey;
    formatDateKey = function (dateObj) {
      var f = formatDateOnly_(dateObj);
      if (f) return f;
      return _formatDateKeyOrig(dateObj);
    };
  }

  if (typeof normalizeImportDate_ !== 'function') {
    normalizeImportDate_ = function (val) {
      if (typeof toCalendarDate_ === 'function') return toCalendarDate_(val);
      if (typeof parseLeaveDate_ === 'function') return parseLeaveDate_(val);
      return null;
    };
  }
  if (typeof sheetDateForWrite_ !== 'function') {
    sheetDateForWrite_ = function (val) {
      if (typeof toSheetDateValue_ === 'function') {
        var v = toSheetDateValue_(val);
        return v === '' ? null : v;
      }
      var d = normalizeImportDate_(val);
      if (!d) return null;
      return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
    };
  }

  Logger.log('ImportDateBridge active: Start/End use DateUtils TZ nudge (min 2h).');
})();
