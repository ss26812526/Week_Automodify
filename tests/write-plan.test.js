(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var W = isNode ? require('../src/write-plan.js') : self.WritePlan;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  var model = [
    { day: 1, weekday: '五', isHoliday: false, isMakeup: false, description: '' },
    { day: 2, weekday: '六', isHoliday: true, isMakeup: false, description: '' },
    { day: 3, weekday: '日', isHoliday: false, isMakeup: true, description: '補行上班' }
  ];

  var cols = { headerRow: 0, dateCol: 0, weekCol: 1, checkCols: [2, 3] };

  function makeTable(dataRows) {
    return [['日期', '星期', '溫濕度', '備註']].concat(dataRows);
  }

  function getVal(result, row, col) {
    var hit = result.writes.filter(function (w) { return w.row === row && w.col === col; });
    return hit.length ? hit[0].value : undefined;
  }

  function assign(a, b) {
    var out = {};
    for (var k in a) out[k] = a[k];
    for (var k2 in b) out[k2] = b[k2];
    return out;
  }

  describe('寫入計畫 › 日期與星期', function () {
    var r = W.buildWritePlan(assign({
      monthModel: model,
      table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, cols));

    it('日期欄依序寫入天數', function () {
      assertEqual(getVal(r, 1, 0), '1');
      assertEqual(getVal(r, 2, 0), '2');
      assertEqual(getVal(r, 3, 0), '3');
    });

    it('星期欄寫入對應星期', function () {
      assertEqual(getVal(r, 1, 1), '五');
      assertEqual(getVal(r, 2, 1), '六');
      assertEqual(getVal(r, 3, 1), '日');
    });
  });

  describe('寫入計畫 › 檢查欄', function () {
    var r = W.buildWritePlan(assign({
      monthModel: model,
      table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, cols));

    it('平日的檢查欄清空', function () {
      assertEqual(getVal(r, 1, 2), '');
      assertEqual(getVal(r, 1, 3), '');
    });

    it('假日的檢查欄填入減號', function () {
      assertEqual(getVal(r, 2, 2), '-');
      assertEqual(getVal(r, 2, 3), '-');
    });

    it('補班日視同平日，檢查欄清空而非填減號', function () {
      assertEqual(getVal(r, 3, 2), '');
      assertEqual(getVal(r, 3, 3), '');
    });
  });

  describe('寫入計畫 › 列摘要', function () {
    var r = W.buildWritePlan(assign({
      monthModel: model,
      table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, cols));

    it('每個資料列都有一筆摘要', function () {
      assertEqual(r.rowSummaries.length, 3);
    });

    it('摘要標示出列的種類', function () {
      assertEqual(r.rowSummaries[0].kind, 'weekday');
      assertEqual(r.rowSummaries[1].kind, 'holiday');
      assertEqual(r.rowSummaries[2].kind, 'makeup');
    });

    it('摘要帶出假日名稱', function () {
      assertEqual(r.rowSummaries[2].description, '補行上班');
    });
  });

  describe('寫入計畫 › 多餘與不足的列', function () {
    it('超出天數的列整列清空並記錄', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['9', '一', 'V', 'V']])
      }, cols));
      assertDeepEqual(r.clearedRows, [4]);
      assertEqual(getVal(r, 4, 0), '');
      assertEqual(getVal(r, 4, 1), '');
      assertEqual(getVal(r, 4, 2), '');
      assertEqual(r.rowSummaries[3].kind, 'cleared');
    });

    it('列數不足時記錄寫不下的日子', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', '', '']])
      }, cols));
      assertDeepEqual(r.missingDays, [2, 3]);
    });

    it('列數剛好時沒有多餘也沒有不足', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
      }, cols));
      assertDeepEqual(r.missingDays, []);
      assertDeepEqual(r.clearedRows, []);
    });
  });

  describe('寫入計畫 › 文字內容警示', function () {
    it('備註欄有真正文字時列入警示', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', 'V', '空調異音已通報'], ['', '', '', ''], ['', '', '', '']])
      }, cols));
      assertEqual(r.contentLoss.length, 1);
      assertEqual(r.contentLoss[0].row, 1);
      assertEqual(r.contentLoss[0].col, 3);
      assertEqual(r.contentLoss[0].text, '空調異音已通報');
    });

    it('勾記符號不列入警示', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', 'V', 'X'], ['', '', '○', '-'], ['', '', '', '']])
      }, cols));
      assertDeepEqual(r.contentLoss, []);
    });

    it('被整列清空的多餘列也會檢查文字', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['9', '一', 'V', '前月殘留紀錄']])
      }, cols));
      assertEqual(r.contentLoss.length, 1);
      assertEqual(r.contentLoss[0].text, '前月殘留紀錄');
    });

    it('日期欄與星期欄的原內容不列入警示', function () {
      var r = W.buildWritePlan(assign({
        monthModel: model,
        table: makeTable([['31', '三', '', ''], ['', '', '', ''], ['', '', '', '']])
      }, cols));
      assertDeepEqual(r.contentLoss, []);
    });
  });
})();
