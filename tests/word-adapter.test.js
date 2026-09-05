(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var D = isNode ? require('../src/word-adapter.js') : self.WordAdapter;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('Word › 文字片段', function () {
    var 段落 = '<w:p><w:r><w:t xml:space="preserve">機房檢查表  </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>115年8月</w:t></w:r></w:p>';

    it('取得段落合併後的完整文字', function () {
      assertEqual(D.mergedText(段落), '機房檢查表  115年8月');
    });

    it('切出兩個片段', function () {
      assertEqual(D.splitSegments(段落).length, 2);
    });

    it('每個片段記錄其在合併文字中的位置', function () {
      var s = D.splitSegments(段落);
      assertEqual(s[0].start, 0);
      assertEqual(s[0].end, 7);
      assertEqual(s[1].start, 7);
      assertEqual(s[1].end, 13);
    });

    it('片段保留原始開始標籤', function () {
      var s = D.splitSegments(段落);
      assertEqual(s[0].openTag, '<w:t xml:space="preserve">');
      assertEqual(s[1].openTag, '<w:t>');
    });

    it('還原跳脫字元', function () {
      assertEqual(D.mergedText('<w:t>A&amp;B</w:t>'), 'A&B');
    });

    it('沒有文字時回傳空字串與空陣列', function () {
      assertEqual(D.mergedText('<w:p></w:p>'), '');
      assertDeepEqual(D.splitSegments('<w:p></w:p>'), []);
    });
  });

  describe('Word › 表格解析', function () {
    var 文件 =
      '<w:body><w:p><w:r><w:t>標題</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr/>' +
      '<w:tr><w:tc><w:p><w:r><w:t>日期</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>星期</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p></w:p></w:tc></w:tr>' +
      '</w:tbl></w:body>';

    it('找得到一個表格', function () {
      assertEqual(D.parseTables(文件).length, 1);
    });

    it('表格有兩列', function () {
      assertEqual(D.parseTables(文件)[0].rows.length, 2);
    });

    it('每列有兩格', function () {
      assertEqual(D.parseTables(文件)[0].rows[0].cells.length, 2);
    });

    it('讀得出儲存格文字', function () {
      var t = D.parseTables(文件)[0];
      assertEqual(t.rows[0].cells[0].text, '日期');
      assertEqual(t.rows[0].cells[1].text, '星期');
      assertEqual(t.rows[1].cells[0].text, '1');
    });

    it('空儲存格為空字串', function () {
      assertEqual(D.parseTables(文件)[0].rows[1].cells[1].text, '');
    });

    it('轉成二維陣列', function () {
      var grid = D.tableToGrid(D.parseTables(文件)[0]);
      assertDeepEqual(grid, [['日期', '星期'], ['1', '']]);
    });

    it('跨多個 run 的儲存格文字會被串接', function () {
      var doc = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>115</w:t></w:r>' +
        '<w:r><w:t>年8月</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
      assertEqual(D.parseTables(doc)[0].rows[0].cells[0].text, '115年8月');
    });

    it('沒有表格時回傳空陣列', function () {
      assertDeepEqual(D.parseTables('<w:body><w:p/></w:body>'), []);
    });
  });

  describe('Word › 讀取真實範例檔', function () {
    if (!isNode) return;
    var fs = require('fs');
    var JSZip = require('jszip');
    var A = require('../src/table-analysis.js');

    it('讀得到一個表格與 32 列', async function () {
      var doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      assertEqual(doc.tables.length, 1);
      assertEqual(doc.grids[0].length, 32, '表頭一列加 31 個資料列');
    });

    it('欄位偵測在真實檔案上成功', async function () {
      var doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      var cols = A.detectColumns(doc.grids[0]);
      assertEqual(cols.headerRow, 0);
      assertEqual(cols.dateCol, 0);
      assertEqual(cols.weekCol, 1);
      assertEqual(cols.checkCols.length, 5);
    });

    it('備註欄的文字讀得到', async function () {
      var doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      assertEqual(doc.grids[0][3][6], '空調異音已通報維護廠商', '第 3 天的備註');
    });
  });

  describe('Word › 範圍改寫', function () {
    it('改寫落在單一片段內的範圍', function () {
      var c = '<w:r><w:t>機房檢查表</w:t></w:r>';
      var out = D.replaceRange(c, D.splitSegments(c), 2, 4, '狀態');
      assertEqual(D.mergedText(out), '機房狀態表');
    });

    it('改寫時保留該片段的開始標籤', function () {
      var c = '<w:r><w:t xml:space="preserve">機房檢查表</w:t></w:r>';
      var out = D.replaceRange(c, D.splitSegments(c), 2, 4, '狀態');
      assertEqual(out.indexOf('xml:space="preserve"') !== -1, true);
    });

    it('改寫跨兩個片段的範圍', function () {
      var c = '<w:r><w:t>115</w:t></w:r><w:r><w:t>年8月</w:t></w:r>';
      var out = D.replaceRange(c, D.splitSegments(c), 0, 6, '115年9月');
      assertEqual(D.mergedText(out), '115年9月');
    });

    it('跨片段改寫後第二個片段只剩範圍外的文字', function () {
      var c = '<w:r><w:t>甲115</w:t></w:r><w:r><w:t>年8月乙</w:t></w:r>';
      var out = D.replaceRange(c, D.splitSegments(c), 1, 7, '115年9月');
      assertEqual(D.mergedText(out), '甲115年9月乙');
    });

    it('新文字中的特殊字元會被跳脫', function () {
      var c = '<w:r><w:t>abc</w:t></w:r>';
      var out = D.replaceRange(c, D.splitSegments(c), 0, 3, 'a<b');
      assertEqual(out.indexOf('&lt;') !== -1, true);
      assertEqual(D.mergedText(out), 'a<b');
    });
  });

  describe('Word › 儲存格寫入', function () {
    it('改寫既有文字', function () {
      var tc = '<w:tc><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '新')), '新');
    });

    it('多個 run 時只留下新值', function () {
      var tc = '<w:tc><w:p><w:r><w:t>甲</w:t></w:r><w:r><w:t>乙</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '丙')), '丙');
    });

    it('清空為空字串', function () {
      var tc = '<w:tc><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '')), '');
    });

    it('原本沒有文字的儲存格能寫入新值', function () {
      var tc = '<w:tc><w:p></w:p></w:tc>';
      var out = D.setCellText(tc, '六');
      assertEqual(D.mergedText(out), '六');
      assertEqual(out.indexOf('<w:tc>') === 0, true, '仍是合法的儲存格');
    });

    it('儲存格屬性完整保留', function () {
      var tc = '<w:tc><w:tcPr><w:tcW w:w="500"/></w:tcPr><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      var out = D.setCellText(tc, '新');
      assertEqual(out.indexOf('<w:tcW w:w="500"/>') !== -1, true);
    });
  });

  describe('Word › 段落年月改寫', function () {
    var doc = '<w:body><w:p><w:r><w:t>機房檢查表  </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>115年8月</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>製表：115年8月</w:t></w:r></w:p></w:body>';

    it('找得到兩個含年月的段落', function () {
      assertEqual(D.findParagraphYearMonth(doc).length, 2);
    });

    it('全部改寫後兩處都變成新年月', function () {
      var targets = D.findParagraphYearMonth(doc);
      var out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, null);
      assertEqual(out.indexOf('115年8月') === -1, true);
      assertEqual((out.match(/115年9月/g) || []).length, 2);
    });

    it('年月完整落在單一 run 時粗體格式得以保留', function () {
      var targets = D.findParagraphYearMonth(doc);
      var out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, null);
      assertEqual(out.indexOf('<w:rPr><w:b/></w:rPr><w:t>115年9月</w:t>') !== -1, true);
    });

    it('只改寫選定的段落', function () {
      var targets = D.findParagraphYearMonth(doc);
      var out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, ['0-0']);
      assertEqual(out.indexOf('製表：115年8月') !== -1, true, '未選取的維持原狀');
    });

    it('年月被切成多個 run 時仍改寫得到', function () {
      var d2 = '<w:body><w:p><w:r><w:t>115</w:t></w:r><w:r><w:t>年8月報表</w:t></w:r></w:p></w:body>';
      var targets = D.findParagraphYearMonth(d2);
      assertEqual(targets.length, 1);
      var out = D.rewriteParagraphYearMonth(d2, targets, 2026, 9, null);
      assertEqual(D.mergedText(out), '115年9月報表');
    });

    it('表格內的年月不會被段落改寫重複處理', function () {
      var d3 = '<w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>115年8月</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>';
      var targets = D.findParagraphYearMonth(d3);
      assertEqual(targets.length, 1, '表格內的段落也算一個目標');
    });
  });

  describe('Word › 完整改寫真實範例檔', function () {
    if (!isNode) return;
    var fs = require('fs');
    var JSZip = require('jszip');
    var A = require('../src/table-analysis.js');
    var M = require('../src/month-model.js');
    var H = require('../src/holidays-tw.js');
    var B = require('../src/holidays-builtin.js');
    var W = require('../src/write-plan.js');

    function 產生改寫結果() {
      return D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip).then(function (doc) {
        var grid = doc.grids[0];
        var cols = A.detectColumns(grid);
        var model = M.buildMonth(2026, 9, H.normalizeYearData(B['2026']));
        var plan = W.buildWritePlan({
          monthModel: model, table: grid, headerRow: cols.headerRow,
          dateCol: cols.dateCol, weekCol: cols.weekCol, checkCols: cols.checkCols
        });
        var xml = D.applyWritesToDocument(doc.documentXml, doc.tables[0], plan.writes);
        var targets = D.findParagraphYearMonth(xml);
        xml = D.rewriteParagraphYearMonth(xml, targets, 2026, 9, null);
        return D.saveDocument(doc.zip, xml, JSZip, 'nodebuffer').then(function (out) {
          return { out: out, plan: plan };
        });
      });
    }

    it('改寫後的檔案仍可被重新讀取', async function () {
      var r = await 產生改寫結果();
      var doc2 = await D.loadDocument(r.out, JSZip);
      assertEqual(doc2.grids[0][0][0], '日期', '表頭未受影響');
    });

    it('日期與星期正確寫入', async function () {
      var r = await 產生改寫結果();
      var g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[1][0], '1');
      assertEqual(g[1][1], '二', '9月1日為星期二');
      assertEqual(g[5][0], '5');
      assertEqual(g[5][1], '六', '9月5日為星期六');
    });

    it('假日的檢查欄填入減號', async function () {
      var r = await 產生改寫結果();
      var g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[5][2], '-', '9月5日為週六');
      assertEqual(g[25][2], '-', '9月25日為中秋節');
    });

    it('平日的檢查欄被清空', async function () {
      var r = await 產生改寫結果();
      var g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[1][2], '', '9月1日為平日');
    });

    it('超出當月天數的第 31 列被整列清空', async function () {
      var r = await 產生改寫結果();
      var g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[31][0], '', '9月只有 30 天');
      assertEqual(g[31][1], '');
    });

    it('標題的年月被改寫且保留粗體', async function () {
      var r = await 產生改寫結果();
      var doc2 = await D.loadDocument(r.out, JSZip);
      assertEqual(doc2.documentXml.indexOf('115年9月') !== -1, true);
      assertEqual(doc2.documentXml.indexOf('115年8月') === -1, true);
      assertEqual(/115年9月<\/w:t><\/w:r>/.test(doc2.documentXml), true, '年月仍在獨立 run 中');
      assertEqual(/<w:b\/><\/w:rPr>[\s\S]*?115年9月/.test(doc2.documentXml) ||
        /<w:b\/><w:bCs\/><\/w:rPr>[\s\S]*?115年9月/.test(doc2.documentXml), true,
        '年月所在的 run 帶有粗體格式');
    });

    it('表格框線設定完全未被更動', async function () {
      var r = await 產生改寫結果();
      var doc2 = await D.loadDocument(r.out, JSZip);
      assertEqual(doc2.documentXml.indexOf('<w:tblBorders>') !== -1, true);
      assertEqual(doc2.documentXml.indexOf('<w:insideH w:val="single"') !== -1, true);
    });
  });
})();
