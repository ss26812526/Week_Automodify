(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var U = isNode ? require('../src/ui-calendar.js') : self.UiCalendar;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual;

  describe('介面 › 預設年月', function () {
    it('取自傳入的日期', function () {
      var r = U.defaultYearMonth(new Date(2026, 8, 30));
      assertEqual(r.year, 2026);
      assertEqual(r.month, 9, '月份為 1 起算');
    });

    it('一月不會變成零月', function () {
      assertEqual(U.defaultYearMonth(new Date(2026, 0, 15)).month, 1);
    });

    it('十二月正確', function () {
      assertEqual(U.defaultYearMonth(new Date(2026, 11, 1)).month, 12);
    });

    it('省略參數時取系統時間且落在合理範圍', function () {
      var r = U.defaultYearMonth();
      assertEqual(r.month >= 1 && r.month <= 12, true);
      assertEqual(r.year >= 2020 && r.year <= 2100, true);
    });
  });

  describe('介面 › 星期索引', function () {
    it('日為 0、六為 6', function () {
      assertEqual(U.weekdayIndex('日'), 0);
      assertEqual(U.weekdayIndex('一'), 1);
      assertEqual(U.weekdayIndex('六'), 6);
    });

    it('無法辨識時回傳 -1', function () {
      assertEqual(U.weekdayIndex('X'), -1);
    });
  });

  describe('介面 › 月曆排列', function () {
    var 模型 = [];
    var 星期 = ['二','三','四','五','六','日','一'];
    for (var d = 1; d <= 30; d++) {
      模型.push({ day: d, weekday: 星期[(d - 1) % 7], isHoliday: false, isMakeup: false, description: '' });
    }

    it('第一週開頭補上空格', function () {
      var weeks = U.buildCalendarWeeks(模型, 2);
      assertEqual(weeks[0][0], null, '星期日位置為空');
      assertEqual(weeks[0][1], null, '星期一位置為空');
      assertEqual(weeks[0][2].day, 1, '星期二為 1 日');
    });

    it('每一週都是七格', function () {
      U.buildCalendarWeeks(模型, 2).forEach(function (w) { assertEqual(w.length, 7); });
    });

    it('最後一天位置正確', function () {
      var weeks = U.buildCalendarWeeks(模型, 2);
      var last = weeks[weeks.length - 1];
      var 有值 = last.filter(function (c) { return c !== null; });
      assertEqual(有值[有值.length - 1].day, 30);
    });

    it('月初為星期日時第一週不補空格', function () {
      var m = [{ day: 1, weekday: '日', isHoliday: true, isMakeup: false, description: '' }];
      assertEqual(U.buildCalendarWeeks(m, 0)[0][0].day, 1);
    });
  });

  describe('介面 › 資料來源標示', function () {
    it('線上取得的說明文字', function () {
      assertEqual(U.sourceLabel('online', 2026), '已取得 2026 年最新假日資料');
    });

    it('內建資料的說明文字', function () {
      assertEqual(U.sourceLabel('builtin', 2026), '使用內建的 2026 年假日資料（無法連線）');
    });

    it('純計算的說明文字要帶警告', function () {
      assertEqual(U.sourceLabel('computed', 2031), '⚠ 沒有 2031 年的假日資料，僅依週六日判斷');
    });
  });
})();
