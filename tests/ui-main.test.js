(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var U = isNode ? require('../src/ui-main.js') : self.UiMain;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('介面 › 輸出檔名', function () {
    it('加上民國年月且月份補零', function () {
      assertEqual(U.outputFilename('機房檢查表.xlsx', 2026, 9), '機房檢查表_115年09月.xlsx');
    });

    it('兩位數月份不再補零', function () {
      assertEqual(U.outputFilename('檢查表.docx', 2026, 12), '檢查表_115年12月.docx');
    });

    it('檔名中含有點號時只取最後一段為副檔名', function () {
      assertEqual(U.outputFilename('機房.狀態.檢查表.xlsx', 2026, 9),
        '機房.狀態.檢查表_115年09月.xlsx');
    });

    it('跨年時民國年正確', function () {
      assertEqual(U.outputFilename('表.xlsx', 2027, 1), '表_116年01月.xlsx');
    });

    it('沒有副檔名時直接附加', function () {
      assertEqual(U.outputFilename('檢查表', 2026, 9), '檢查表_115年09月');
    });
  });

  describe('介面 › 格式判斷', function () {
    it('辨識 xlsx 與 docx', function () {
      assertEqual(U.detectFormat('a.xlsx'), 'xlsx');
      assertEqual(U.detectFormat('a.docx'), 'docx');
    });

    it('大寫副檔名也辨識得出', function () {
      assertEqual(U.detectFormat('A.XLSX'), 'xlsx');
    });

    it('舊格式獨立標示以便給出明確提示', function () {
      assertEqual(U.detectFormat('a.xls'), 'legacy');
      assertEqual(U.detectFormat('a.doc'), 'legacy');
    });

    it('其他格式為未知', function () {
      assertEqual(U.detectFormat('a.pdf'), 'unknown');
      assertEqual(U.detectFormat('a'), 'unknown');
    });
  });

  describe('介面 › 欄位記憶鍵值', function () {
    it('相同表頭產生相同鍵值', function () {
      var a = U.storageKey(['日期', '星期', '溫濕度']);
      var b = U.storageKey(['日期', '星期', '溫濕度']);
      assertEqual(a, b);
    });

    it('不同表頭產生不同鍵值', function () {
      var a = U.storageKey(['日期', '星期', '溫濕度']);
      var b = U.storageKey(['日期', '星期', 'UPS']);
      assertEqual(a === b, false);
    });

    it('鍵值有固定前綴便於辨識', function () {
      assertEqual(U.storageKey(['日期']).indexOf('week-automodify:') === 0, true);
    });
  });

  describe('介面 › 由手動指定組出欄位設定', function () {
    it('其餘欄位皆列為檢查欄', function () {
      var r = U.columnsFrom(['日期', '星期', '溫濕度', '備註'], 2, 0, 1);
      assertEqual(r.headerRow, 2);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
      assertDeepEqual(r.checkCols, [2, 3]);
    });

    it('日期與星期不在最前面兩欄也正確', function () {
      var r = U.columnsFrom(['項次', '日期', '星期', '溫濕度'], 0, 1, 2);
      assertDeepEqual(r.checkCols, [0, 3]);
    });
  });
})();
