(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var Y = isNode ? require('../src/year-month-text.js') : self.YearMonthText;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('年月文字 › 偵測', function () {
    it('找得到民國年月', function () {
      var m = Y.findYearMonth('機房狀態檢查表 115年8月');
      assertEqual(m.length, 1);
      assertEqual(m[0].era, 'roc');
      assertEqual(m[0].year, 2026);
      assertEqual(m[0].month, 8);
      assertEqual(m[0].raw, '115年8月');
    });

    it('找得到西元年月', function () {
      var m = Y.findYearMonth('2026年08月 檢查紀錄');
      assertEqual(m.length, 1);
      assertEqual(m[0].era, 'ad');
      assertEqual(m[0].year, 2026);
      assertEqual(m[0].month, 8);
    });

    it('一段文字中的多處年月全部找得到', function () {
      var m = Y.findYearMonth('115年8月報表　　製表日期：115年9月');
      assertEqual(m.length, 2);
      assertEqual(m[0].month, 8);
      assertEqual(m[1].month, 9);
    });

    it('沒有年月時回傳空陣列', function () {
      assertDeepEqual(Y.findYearMonth('機房狀態檢查表'), []);
    });

    it('年份 200 視為民國', function () {
      assertEqual(Y.findYearMonth('200年1月')[0].era, 'roc');
    });

    it('年份 201 視為西元', function () {
      assertEqual(Y.findYearMonth('201年1月')[0].era, 'ad');
    });

    it('記錄下原文中的空白', function () {
      var m = Y.findYearMonth('115 年 8 月');
      assertEqual(m[0].raw, '115 年 8 月');
    });
  });

  describe('年月文字 › 改寫', function () {
    it('民國年改寫後仍為民國年', function () {
      var m = Y.findYearMonth('115年8月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '115年9月');
    });

    it('跨年時民國年份跟著進位', function () {
      var m = Y.findYearMonth('114年12月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 1), '115年1月');
    });

    it('西元年改寫後仍為西元年', function () {
      var m = Y.findYearMonth('2025年12月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 1), '2026年1月');
    });

    it('原本月份補零則改寫後也補零', function () {
      var m = Y.findYearMonth('2026年08月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '2026年09月');
    });

    it('原本月份未補零則改寫後也不補零', function () {
      var m = Y.findYearMonth('115年8月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 12), '115年12月');
    });

    it('原文的空白完整沿用', function () {
      var m = Y.findYearMonth('115 年 8 月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '115 年 9 月');
    });
  });

  describe('年月文字 › 整段套用', function () {
    it('多處年月可一次全部改寫', function () {
      var text = '115年8月報表　　製表：115年8月';
      var m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 9), '115年9月報表　　製表：115年9月');
    });

    it('可只改寫選定的其中一處', function () {
      var text = '115年8月報表　　製表：115年8月';
      var m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 9, [0]), '115年9月報表　　製表：115年8月');
    });

    it('改寫後長度變動不影響其他處的位置', function () {
      var text = '115年9月　115年9月';
      var m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 12), '115年12月　115年12月');
    });

    it('沒有匹配時原文不變', function () {
      assertEqual(Y.applyRewrites('機房檢查表', [], 2026, 9), '機房檢查表');
    });
  });
})();
