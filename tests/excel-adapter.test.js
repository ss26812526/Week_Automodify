(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var E = isNode ? require('../src/excel-adapter.js') : self.ExcelAdapter;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('Excel › 欄名轉換', function () {
    it('單字母欄名', function () {
      assertEqual(E.colToIndex('A'), 0);
      assertEqual(E.colToIndex('B'), 1);
      assertEqual(E.colToIndex('Z'), 25);
    });

    it('雙字母欄名', function () {
      assertEqual(E.colToIndex('AA'), 26);
      assertEqual(E.colToIndex('AB'), 27);
      assertEqual(E.colToIndex('BA'), 52);
    });

    it('索引轉回欄名', function () {
      assertEqual(E.indexToCol(0), 'A');
      assertEqual(E.indexToCol(25), 'Z');
      assertEqual(E.indexToCol(26), 'AA');
      assertEqual(E.indexToCol(27), 'AB');
    });

    it('來回轉換一致', function () {
      for (var i = 0; i < 100; i++) assertEqual(E.colToIndex(E.indexToCol(i)), i);
    });

    it('儲存格參照拆解為零起算的列與欄', function () {
      assertDeepEqual(E.parseCellRef('A1'), { row: 0, col: 0 });
      assertDeepEqual(E.parseCellRef('B5'), { row: 4, col: 1 });
      assertDeepEqual(E.parseCellRef('AA10'), { row: 9, col: 26 });
    });
  });

  describe('Excel › 共用字串解析', function () {
    it('取出所有字串', function () {
      var xml = '<sst count="3" uniqueCount="3"><si><t>日期</t></si>' +
        '<si><t>星期</t></si><si><t>備註</t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['日期', '星期', '備註']);
    });

    it('空字串項目保留為空字串', function () {
      var xml = '<sst><si><t>甲</t></si><si><t></t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['甲', '']);
    });

    it('還原跳脫字元', function () {
      var xml = '<sst><si><t>A&amp;B</t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['A&B']);
    });

    it('處理被切成多段的富文字', function () {
      var xml = '<sst><si><r><t>115年</t></r><r><t>8月</t></r></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['115年8月']);
    });

    it('沒有共用字串時回傳空陣列', function () {
      assertDeepEqual(E.parseSharedStrings('<sst/>'), []);
    });
  });

  describe('Excel › 工作表解析', function () {
    var shared = ['標題', '日期', '星期', '備註', '六'];
    var sheetXml =
      '<worksheet><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' +
      '<row r="3"><c r="A3" s="1" t="s"><v>1</v></c>' +
      '<c r="B3" s="1" t="s"><v>2</v></c><c r="C3" s="1" t="s"><v>3</v></c></row>' +
      '<row r="4"><c r="A4" s="1"><v>1</v></c>' +
      '<c r="B4" s="1" t="s"><v>4</v></c><c r="C4" s="1"/></row>' +
      '</sheetData></worksheet>';

    it('依 r 屬性定位，跳過的列成為空列', function () {
      var r = E.parseSheet(sheetXml, shared);
      assertEqual(r.table.length, 4, '應有第 1 至第 4 列');
      assertDeepEqual(r.table[1], [], '第 2 列在 XML 中不存在，應為空列');
    });

    it('共用字串型的儲存格取出文字而非索引', function () {
      var r = E.parseSheet(sheetXml, shared);
      assertEqual(r.table[2][0], '日期');
      assertEqual(r.table[2][1], '星期');
    });

    it('數字型儲存格取出數字文字', function () {
      assertEqual(E.parseSheet(sheetXml, shared).table[3][0], '1');
    });

    it('自閉合的空儲存格為空字串', function () {
      assertEqual(E.parseSheet(sheetXml, shared).table[3][2], '');
    });

    it('保留每一格原本的樣式屬性', function () {
      var r = E.parseSheet(sheetXml, shared);
      var b4 = r.cells.filter(function (c) { return c.ref === 'B4'; })[0];
      assertEqual(b4.styleAttr, ' s="1"');
      assertEqual(b4.row, 3);
      assertEqual(b4.col, 1);
    });

    it('行內字串型的儲存格也讀得到', function () {
      var xml = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>行內</t></is></c></row></sheetData>';
      assertEqual(E.parseSheet(xml, []).table[0][0], '行內');
    });
  });

  describe('Excel › 讀取真實範例檔', function () {
    if (!isNode) return;
    var fs = require('fs');
    var JSZip = require('jszip');
    var A = require('../src/table-analysis.js');

    it('讀得出表格內容', async function () {
      var buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      var wb = await E.loadWorkbook(buf, JSZip);
      assertEqual(wb.table[0][0], '機房狀態月檢查表  115年8月', '第 1 列為標題');
      assertEqual(wb.table[2][0], '日期', '第 3 列為表頭');
      assertEqual(wb.table[2][1], '星期');
      assertEqual(wb.table[3][0], '1', '第 4 列為第一天');
    });

    it('欄位偵測能在真實檔案上成功', async function () {
      var buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      var wb = await E.loadWorkbook(buf, JSZip);
      var cols = A.detectColumns(wb.table);
      assertEqual(cols.headerRow, 2);
      assertEqual(cols.dateCol, 0);
      assertEqual(cols.weekCol, 1);
      assertEqual(cols.checkCols.length, 5);
    });

    it('備註欄的文字讀得到', async function () {
      var buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      var wb = await E.loadWorkbook(buf, JSZip);
      assertEqual(wb.table[5][6], '空調異音已通報維護廠商', '第 3 天的備註');
    });
  });

  describe('Excel › 儲存格 XML 產生', function () {
    it('有值時產生行內字串', function () {
      assertEqual(E.buildCellXml('B4', ' s="1"', '六'),
        '<c r="B4" s="1" t="inlineStr"><is><t>六</t></is></c>');
    });

    it('空值時產生自閉合標籤並保留樣式', function () {
      assertEqual(E.buildCellXml('C4', ' s="1"', ''), '<c r="C4" s="1"/>');
    });

    it('沒有樣式屬性時也能產生', function () {
      assertEqual(E.buildCellXml('A1', '', 'X'),
        '<c r="A1" t="inlineStr"><is><t>X</t></is></c>');
    });

    it('值中的特殊字元會被跳脫', function () {
      assertEqual(E.buildCellXml('A1', '', 'a<b&c'),
        '<c r="A1" t="inlineStr"><is><t>a&lt;b&amp;c</t></is></c>');
    });

    it('減號可正常寫入', function () {
      assertEqual(E.buildCellXml('C5', ' s="2"', '-'),
        '<c r="C5" s="2" t="inlineStr"><is><t>-</t></is></c>');
    });
  });

  describe('Excel › 套用寫入', function () {
    var writeShared = ['日期', '星期', '六'];
    var writeSheet =
      '<worksheet><sheetData>' +
      '<row r="3"><c r="A3" s="1" t="s"><v>0</v></c><c r="B3" s="1" t="s"><v>1</v></c></row>' +
      '<row r="4"><c r="A4" s="1"><v>1</v></c><c r="B4" s="1" t="s"><v>2</v></c>' +
      '<c r="C4" s="1"><v>99</v></c></row>' +
      '</sheetData></worksheet>';

    it('改寫後的值讀得回來', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [{ row: 3, col: 1, value: '日' }], parsed.cells);
      var reread = E.parseSheet(r.xml, writeShared);
      assertEqual(reread.table[3][1], '日');
    });

    it('清空後該格為空字串', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [{ row: 3, col: 2, value: '' }], parsed.cells);
      assertEqual(E.parseSheet(r.xml, writeShared).table[3][2], '');
    });

    it('樣式屬性完整保留', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [{ row: 3, col: 2, value: '' }], parsed.cells);
      assertEqual(r.xml.indexOf('<c r="C4" s="1"/>') !== -1, true, '應保留 s="1"');
    });

    it('未被寫入的儲存格原封不動', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [{ row: 3, col: 1, value: '日' }], parsed.cells);
      assertEqual(r.xml.indexOf('<c r="A3" s="1" t="s"><v>0</v></c>') !== -1, true);
    });

    it('目標儲存格不存在時列入略過清單', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [{ row: 3, col: 9, value: 'X' }], parsed.cells);
      assertEqual(r.skipped.length, 1);
      assertEqual(r.skipped[0].col, 9);
    });

    it('一次寫入多格皆生效', function () {
      var parsed = E.parseSheet(writeSheet, writeShared);
      var r = E.applyWritesToSheet(writeSheet, [
        { row: 3, col: 0, value: '5' },
        { row: 3, col: 1, value: '三' },
        { row: 3, col: 2, value: '-' }
      ], parsed.cells);
      var reread = E.parseSheet(r.xml, writeShared);
      assertEqual(reread.table[3][0], '5');
      assertEqual(reread.table[3][1], '三');
      assertEqual(reread.table[3][2], '-');
    });
  });

  describe('Excel › 共用字串中的年月改寫', function () {
    var ymXml = '<?xml version="1.0"?><sst count="3" uniqueCount="3">' +
      '<si><t>機房狀態月檢查表  115年8月</t></si>' +
      '<si><t>日期</t></si><si><t>製表：115年8月</t></si></sst>';
    var ymShared = ['機房狀態月檢查表  115年8月', '日期', '製表：115年8月'];

    it('找得到所有含年月的字串', function () {
      var targets = E.findYearMonthTargets(ymShared);
      assertEqual(targets.length, 2);
      assertEqual(targets[0].index, 0);
      assertEqual(targets[1].index, 2);
    });

    it('沒有年月的字串不列入', function () {
      var targets = E.findYearMonthTargets(['日期', '星期']);
      assertEqual(targets.length, 0);
    });

    it('全部改寫後兩處都變成新年月', function () {
      var targets = E.findYearMonthTargets(ymShared);
      var out = E.rewriteSharedStrings(ymXml, ymShared, targets, 2026, 9, null);
      assertEqual(out.indexOf('115年9月') !== -1, true);
      assertEqual(out.indexOf('115年8月') === -1, true, '舊年月應全部被取代');
    });

    it('只改寫選定的其中一處', function () {
      var targets = E.findYearMonthTargets(ymShared);
      var out = E.rewriteSharedStrings(ymXml, ymShared, targets, 2026, 9, ['0-0']);
      assertEqual(out.indexOf('機房狀態月檢查表  115年9月') !== -1, true);
      assertEqual(out.indexOf('製表：115年8月') !== -1, true, '未選取的應維持原狀');
    });

    it('count 與 uniqueCount 不被更動', function () {
      var targets = E.findYearMonthTargets(ymShared);
      var out = E.rewriteSharedStrings(ymXml, ymShared, targets, 2026, 9, null);
      assertEqual(out.indexOf('count="3" uniqueCount="3"') !== -1, true);
    });

    it('改寫後仍解析得回正確字串', function () {
      var targets = E.findYearMonthTargets(ymShared);
      var out = E.rewriteSharedStrings(ymXml, ymShared, targets, 2026, 12, null);
      assertEqual(E.parseSharedStrings(out)[0], '機房狀態月檢查表  115年12月');
    });
  });

  describe('Excel › 完整改寫真實範例檔', function () {
    if (!isNode) return;
    var fs = require('fs');
    var JSZip = require('jszip');
    var A = require('../src/table-analysis.js');
    var M = require('../src/month-model.js');
    var H = require('../src/holidays-tw.js');
    var B = require('../src/holidays-builtin.js');
    var W = require('../src/write-plan.js');

    function generateResult() {
      var buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      return E.loadWorkbook(buf, JSZip).then(function (wb) {
        var cols = A.detectColumns(wb.table);
        var model = M.buildMonth(2026, 9, H.normalizeYearData(B['2026']));
        var plan = W.buildWritePlan({
          monthModel: model, table: wb.table, headerRow: cols.headerRow,
          dateCol: cols.dateCol, weekCol: cols.weekCol, checkCols: cols.checkCols
        });
        var applied = E.applyWritesToSheet(wb.sheetXml, plan.writes, wb.cells);
        var targets = E.findYearMonthTargets(wb.sharedStrings);
        var newSs = E.rewriteSharedStrings(wb.sharedStringsXml, wb.sharedStrings, targets, 2026, 9, null);
        return E.saveWorkbook(wb.zip, wb.sheetPath, applied.xml, newSs, JSZip, 'nodebuffer').then(function (out) {
          return { out: out, plan: plan, applied: applied };
        });
      });
    }

    it('改寫後的檔案仍可被重新讀取', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[2][0], '日期', '表頭未受影響');
    });

    it('日期與星期正確寫入', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[3][0], '1', '9月1日');
      assertEqual(wb2.table[3][1], '二', '9月1日為星期二');
      assertEqual(wb2.table[7][0], '5');
      assertEqual(wb2.table[7][1], '六', '9月5日為星期六');
    });

    it('假日的檢查欄填入減號', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[7][2], '-', '9月5日為週六');
      assertEqual(wb2.table[27][2], '-', '9月25日為中秋節');
    });

    it('平日的檢查欄被清空', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[3][2], '', '9月1日為平日');
    });

    it('超出當月天數的第 31 列被整列清空', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[33][0], '', '9月只有 30 天，第 31 列應清空');
      assertEqual(wb2.table[33][1], '');
    });

    it('標題的年月被改寫', async function () {
      var r = await generateResult();
      var wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[0][0], '機房狀態月檢查表  115年9月');
    });

    it('備註欄的文字被列入警示', async function () {
      var r = await generateResult();
      assertEqual(r.plan.contentLoss.length, 1);
      assertEqual(r.plan.contentLoss[0].text, '空調異音已通報維護廠商');
    });

    it('沒有任何儲存格被略過', async function () {
      var r = await generateResult();
      assertEqual(r.applied.skipped.length, 0);
    });

    it('框線樣式檔完全未被更動', async function () {
      var r = await generateResult();
      var origZip = await JSZip.loadAsync(fs.readFileSync('samples/機房月檢查表-範例.xlsx'));
      var newZip = await JSZip.loadAsync(r.out);
      var a = await origZip.file('xl/styles.xml').async('string');
      var b = await newZip.file('xl/styles.xml').async('string');
      assertEqual(a === b, true, 'styles.xml 必須逐字相同');
    });
  });
})();
