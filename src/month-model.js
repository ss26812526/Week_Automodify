// 月份模型：把日曆骨架與假日資料合併成介面與寫入計畫共用的結構
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./calendar-core.js'));
  } else {
    root.MonthModel = factory(root.CalendarCore);
  }
})(typeof self !== 'undefined' ? self : this, function (CalendarCore) {

  function buildMonth(year, month, yearData) {
    var data = yearData || {};
    return CalendarCore.buildMonthSkeleton(year, month).map(function (row) {
      var key = CalendarCore.formatDateKey(year, month, row.day);
      var found = data[key];

      if (!found) {
        var isWeekend = row.weekday === '六' || row.weekday === '日';
        return {
          day: row.day,
          weekday: row.weekday,
          isHoliday: isWeekend,
          isMakeup: false,
          description: ''
        };
      }

      var weekday = found.week || row.weekday;
      var isWeekend2 = weekday === '六' || weekday === '日';
      return {
        day: row.day,
        weekday: weekday,
        isHoliday: found.isHoliday,
        isMakeup: isWeekend2 && !found.isHoliday,
        description: found.description || ''
      };
    });
  }

  return { buildMonth: buildMonth };
});
