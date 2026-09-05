// 假日資料層：線上取得 → 內建資料 → 純計算，三層遞降
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./calendar-core.js'));
  } else {
    root.HolidaysTW = factory(root.CalendarCore);
  }
})(typeof self !== 'undefined' ? self : this, function (CalendarCore) {
  var BASE_URL = 'https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/';
  var DEFAULT_TIMEOUT_MS = 4000;

  function buildUrl(year) {
    return BASE_URL + year + '.json';
  }

  function normalizeYearData(raw) {
    var out = {};
    (raw || []).forEach(function (row) {
      out[row.date] = {
        week: row.week,
        isHoliday: !!row.isHoliday,
        description: row.description || ''
      };
    });
    return out;
  }

  function computeYearData(year) {
    var out = {};
    for (var month = 1; month <= 12; month++) {
      var total = CalendarCore.daysInMonth(year, month);
      for (var day = 1; day <= total; day++) {
        var week = CalendarCore.weekdayOf(year, month, day);
        out[CalendarCore.formatDateKey(year, month, day)] = {
          week: week,
          isHoliday: week === '六' || week === '日',
          description: ''
        };
      }
    }
    return out;
  }

  function withTimeout(promise, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var timer = setTimeout(function () { reject(new Error('連線逾時')); }, timeoutMs);
      promise.then(
        function (v) { clearTimeout(timer); resolve(v); },
        function (e) { clearTimeout(timer); reject(e); }
      );
    });
  }

  function fetchOnlineYear(year, fetchImpl, timeoutMs) {
    if (!fetchImpl) return Promise.reject(new Error('此環境沒有可用的 fetch'));
    var request = Promise.resolve(fetchImpl(buildUrl(year))).then(function (res) {
      if (!res || !res.ok) throw new Error('回應狀態異常：' + (res ? res.status : '無回應'));
      return res.json();
    }).then(normalizeYearData);
    return withTimeout(request, timeoutMs || DEFAULT_TIMEOUT_MS);
  }

  function getYearData(year, options) {
    var opts = options || {};
    var fetchImpl = opts.fetchImpl !== undefined
      ? opts.fetchImpl
      : (typeof fetch !== 'undefined' ? fetch : null);
    var builtin = opts.builtin || {};

    return fetchOnlineYear(year, fetchImpl, opts.timeoutMs)
      .then(function (data) {
        return { source: 'online', data: data };
      })
      .catch(function () {
        if (builtin[year]) {
          return { source: 'builtin', data: normalizeYearData(builtin[year]) };
        }
        return { source: 'computed', data: computeYearData(year) };
      });
  }

  return {
    BASE_URL: BASE_URL,
    buildUrl: buildUrl,
    normalizeYearData: normalizeYearData,
    computeYearData: computeYearData,
    fetchOnlineYear: fetchOnlineYear,
    getYearData: getYearData
  };
});
