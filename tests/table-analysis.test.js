(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var A = isNode ? require('../src/table-analysis.js') : self.TableAnalysis;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('表格分析 › 記號判定', function () {
    it('空字串視為無資訊', function () {
      assertEqual(A.isSymbolOnly(''), true);
    });

    it('純空白視為無資訊', function () {
      assertEqual(A.isSymbolOnly('   '), true);
    });

    it('勾記符號視為無資訊', function () {
      ['V', 'v', 'X', 'x', '○', '✓', '✗', '-', '－'].forEach(function (s) {
        assertEqual(A.isSymbolOnly(s), true, s + ' 應視為記號');
      });
    });

    it('記號前後有空白仍視為無資訊', function () {
      assertEqual(A.isSymbolOnly(' V '), true);
    });

    it('真正的文字不視為記號', function () {
      assertEqual(A.isSymbolOnly('空調異音已通報'), false);
    });

    it('多個字元的內容不視為記號', function () {
      assertEqual(A.isSymbolOnly('VV'), false);
      assertEqual(A.isSymbolOnly('OK'), false);
    });

    it('null 與 undefined 視為無資訊', function () {
      assertEqual(A.isSymbolOnly(null), true);
      assertEqual(A.isSymbolOnly(undefined), true);
    });
  });

  describe('表格分析 › 欄位偵測', function () {
    var standardTable = [
      ['機房狀態月檢查表  115年8月', '', '', '', '', ''],
      [],
      ['日期', '星期', '溫濕度', 'UPS', '消防', '備註'],
      ['1', '六', 'V', 'V', 'V', '']
    ];

    it('找得到表頭列與日期星期欄', function () {
      var r = A.detectColumns(standardTable);
      assertEqual(r.headerRow, 2);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
    });

    it('其餘欄位皆列為檢查欄', function () {
      assertDeepEqual(A.detectColumns(standardTable).checkCols, [2, 3, 4, 5]);
    });

    it('日期欄不在第一欄也找得到', function () {
      var t = [['項次', '日期', '星期', '溫濕度']];
      var r = A.detectColumns(t);
      assertEqual(r.dateCol, 1);
      assertEqual(r.weekCol, 2);
      assertDeepEqual(r.checkCols, [0, 3]);
    });

    it('星期欄寫成「週」也找得到', function () {
      var r = A.detectColumns([['日期', '週', '項目']]);
      assertEqual(r.weekCol, 1);
    });

    it('日期欄只寫「日」也找得到', function () {
      var r = A.detectColumns([['日', '星期', '項目']]);
      assertEqual(r.dateCol, 0);
    });

    it('欄位文字前後有空白仍找得到', function () {
      var r = A.detectColumns([[' 日期 ', ' 星期 ', '項目']]);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
    });

    it('只有日期欄沒有星期欄時視為偵測失敗', function () {
      assertEqual(A.detectColumns([['日期', '溫濕度', 'UPS']]), null);
    });

    it('完全沒有表頭時視為偵測失敗', function () {
      assertEqual(A.detectColumns([['1', '2', '3'], ['4', '5', '6']]), null);
    });

    it('表頭在掃描範圍外時視為偵測失敗', function () {
      var t = [];
      for (var i = 0; i < 12; i++) t.push(['', '', '']);
      t.push(['日期', '星期', '項目']);
      assertEqual(A.detectColumns(t, 10), null);
    });

    it('加大掃描列數後就找得到', function () {
      var t = [];
      for (var i = 0; i < 12; i++) t.push(['', '', '']);
      t.push(['日期', '星期', '項目']);
      assertEqual(A.detectColumns(t, 20).headerRow, 12);
    });
  });
})();
