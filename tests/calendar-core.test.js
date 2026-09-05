(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var C = isNode ? require('../src/calendar-core.js') : self.CalendarCore;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('日曆核心 › 月份天數', function () {
    it('平年二月為 28 天', function () {
      assertEqual(C.daysInMonth(2026, 2), 28);
    });

    it('閏年二月為 29 天', function () {
      assertEqual(C.daysInMonth(2024, 2), 29);
    });

    it('世紀年能被 400 整除仍為閏年', function () {
      assertEqual(C.daysInMonth(2000, 2), 29);
    });

    it('世紀年不能被 400 整除則非閏年', function () {
      assertEqual(C.daysInMonth(1900, 2), 28);
    });

    it('小月為 30 天', function () {
      assertEqual(C.daysInMonth(2026, 9), 30);
      assertEqual(C.daysInMonth(2026, 4), 30);
    });

    it('大月為 31 天', function () {
      assertEqual(C.daysInMonth(2026, 1), 31);
      assertEqual(C.daysInMonth(2026, 12), 31);
    });
  });

  describe('日曆核心 › 星期推算', function () {
    it('2026年9月1日為星期二', function () {
      assertEqual(C.weekdayOf(2026, 9, 1), '二');
    });

    it('2026年1月1日為星期四', function () {
      assertEqual(C.weekdayOf(2026, 1, 1), '四');
    });

    it('月初落在星期日', function () {
      assertEqual(C.weekdayOf(2026, 3, 1), '日');
    });

    it('月初落在星期一', function () {
      assertEqual(C.weekdayOf(2026, 6, 1), '一');
    });

    it('跨年後的隔年一月一日星期正確', function () {
      assertEqual(C.weekdayOf(2027, 1, 1), '五');
    });

    it('閏年二月二十九日星期正確', function () {
      assertEqual(C.weekdayOf(2024, 2, 29), '四');
    });
  });

  describe('日曆核心 › 月份骨架', function () {
    it('小月產生 30 筆', function () {
      assertEqual(C.buildMonthSkeleton(2026, 9).length, 30);
    });

    it('平年二月產生 28 筆', function () {
      assertEqual(C.buildMonthSkeleton(2026, 2).length, 28);
    });

    it('第一筆與最後一筆內容正確', function () {
      var rows = C.buildMonthSkeleton(2026, 9);
      assertDeepEqual(rows[0], { day: 1, weekday: '二' });
      assertDeepEqual(rows[29], { day: 30, weekday: '三' });
    });
  });

  describe('日曆核心 › 日期鍵值', function () {
    it('月與日均補零至兩位', function () {
      assertEqual(C.formatDateKey(2026, 9, 5), '20260905');
    });

    it('兩位數的月與日不再補零', function () {
      assertEqual(C.formatDateKey(2026, 12, 25), '20261225');
    });
  });
})();
