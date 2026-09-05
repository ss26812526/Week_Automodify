// 日曆核心：零依賴純函式，只處理天數與星期
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CalendarCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];

  function isLeapYear(year) {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  }

  function daysInMonth(year, month) {
    var table = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (month === 2 && isLeapYear(year)) return 29;
    return table[month - 1];
  }

  function weekdayOf(year, month, day) {
    var d = new Date(Date.UTC(year, month - 1, day));
    return WEEKDAY_NAMES[d.getUTCDay()];
  }

  function buildMonthSkeleton(year, month) {
    var total = daysInMonth(year, month);
    var rows = [];
    for (var day = 1; day <= total; day++) {
      rows.push({ day: day, weekday: weekdayOf(year, month, day) });
    }
    return rows;
  }

  function pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  function formatDateKey(year, month, day) {
    return String(year) + pad2(month) + pad2(day);
  }

  return {
    WEEKDAY_NAMES: WEEKDAY_NAMES,
    isLeapYear: isLeapYear,
    daysInMonth: daysInMonth,
    weekdayOf: weekdayOf,
    buildMonthSkeleton: buildMonthSkeleton,
    formatDateKey: formatDateKey
  };
});
