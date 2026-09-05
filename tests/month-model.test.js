(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var M = isNode ? require('../src/month-model.js') : self.MonthModel;
  var H = isNode ? require('../src/holidays-tw.js') : self.HolidaysTW;
  var B = isNode ? require('../src/holidays-builtin.js') : self.HolidaysBuiltin;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual;

  var y2026 = H.normalizeYearData(B['2026']);
  var y2025 = H.normalizeYearData(B['2025']);

  describe('月份模型', function () {
    it('小月產生 30 筆', function () {
      assertEqual(M.buildMonth(2026, 9, y2026).length, 30);
    });

    it('平日不是假日也不是補班日', function () {
      var d = M.buildMonth(2026, 9, y2026)[0];
      assertEqual(d.day, 1);
      assertEqual(d.weekday, '二');
      assertEqual(d.isHoliday, false);
      assertEqual(d.isMakeup, false);
    });

    it('週六判為假日', function () {
      var d = M.buildMonth(2026, 9, y2026)[4];
      assertEqual(d.day, 5);
      assertEqual(d.weekday, '六');
      assertEqual(d.isHoliday, true);
    });

    it('落在平日的國定假日判為假日並帶出名稱', function () {
      var d = M.buildMonth(2026, 9, y2026)[24];
      assertEqual(d.day, 25);
      assertEqual(d.weekday, '五');
      assertEqual(d.isHoliday, true);
      assertEqual(d.description, '中秋節');
    });

    it('教師節判為假日', function () {
      var d = M.buildMonth(2026, 9, y2026)[27];
      assertEqual(d.day, 28);
      assertEqual(d.isHoliday, true);
    });

    it('補班日判為非假日且標記為補班', function () {
      var feb = M.buildMonth(2025, 2, y2025);
      var d = feb[7];
      assertEqual(d.day, 8);
      assertEqual(d.weekday, '六');
      assertEqual(d.isHoliday, false);
      assertEqual(d.isMakeup, true);
    });

    it('缺少該日資料時退回以六日判斷', function () {
      var sep = M.buildMonth(2026, 9, {});
      assertEqual(sep[4].isHoliday, true, '9月5日為週六');
      assertEqual(sep[0].isHoliday, false, '9月1日為週二');
      assertEqual(sep[24].isHoliday, false, '無資料時中秋節無法辨識');
    });

    it('閏年二月產生 29 筆', function () {
      assertEqual(M.buildMonth(2024, 2, {}).length, 29);
    });
  });
})();
