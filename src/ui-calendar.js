// 介面：年月選擇與月曆呈現
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.UiCalendar = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var WEEK_ORDER = ['日', '一', '二', '三', '四', '五', '六'];

  function defaultYearMonth(now) {
    var d = now || new Date();
    return { year: d.getFullYear(), month: d.getMonth() + 1 };
  }

  function weekdayIndex(weekday) {
    return WEEK_ORDER.indexOf(weekday);
  }

  function buildCalendarWeeks(monthModel, firstWeekdayIndex) {
    var weeks = [];
    var week = [];
    for (var i = 0; i < firstWeekdayIndex; i++) week.push(null);

    monthModel.forEach(function (day) {
      week.push(day);
      if (week.length === 7) { weeks.push(week); week = []; }
    });

    if (week.length > 0) {
      while (week.length < 7) week.push(null);
      weeks.push(week);
    }
    return weeks;
  }

  function sourceLabel(source, year) {
    if (source === 'online') return '已取得 ' + year + ' 年最新假日資料';
    if (source === 'builtin') return '使用內建的 ' + year + ' 年假日資料（無法連線）';
    return '⚠ 沒有 ' + year + ' 年的假日資料，僅依週六日判斷';
  }

  function renderCalendar(monthModel, container, onToggle) {
    container.innerHTML = '';
    if (monthModel.length === 0) return;

    var first = weekdayIndex(monthModel[0].weekday);
    var weeks = buildCalendarWeeks(monthModel, first < 0 ? 0 : first);

    var table = document.createElement('table');
    table.className = 'calendar';

    var thead = document.createElement('thead');
    var hrow = document.createElement('tr');
    WEEK_ORDER.forEach(function (w) {
      var th = document.createElement('th');
      th.textContent = w;
      hrow.appendChild(th);
    });
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = document.createElement('tbody');
    weeks.forEach(function (week) {
      var tr = document.createElement('tr');
      week.forEach(function (day) {
        var td = document.createElement('td');
        if (day) {
          td.textContent = day.day;
          td.title = day.description || '';
          if (day.isHoliday) td.className = 'holiday';
          else if (day.isMakeup) td.className = 'makeup';
          if (day.description) td.classList.add('named');
          td.addEventListener('click', function () { onToggle(day); });
        } else {
          td.className = 'blank';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    container.appendChild(table);
  }

  return {
    WEEK_ORDER: WEEK_ORDER,
    defaultYearMonth: defaultYearMonth,
    weekdayIndex: weekdayIndex,
    buildCalendarWeeks: buildCalendarWeeks,
    sourceLabel: sourceLabel,
    renderCalendar: renderCalendar
  };
});
