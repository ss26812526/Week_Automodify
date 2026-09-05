// 表格分析：欄位偵測與記號判定，Excel 與 Word 共用
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TableAnalysis = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var SYMBOLS = ['V', 'v', 'X', 'x', '○', '●', '◯', '✓', '✔', '✗', '✘', '-', '－', 'ˇ'];

  function isSymbolOnly(text) {
    if (text === null || text === undefined) return true;
    var t = String(text).trim();
    if (t === '') return true;
    return SYMBOLS.indexOf(t) !== -1;
  }

  function normalizeHeader(text) {
    return String(text === null || text === undefined ? '' : text).trim();
  }

  function findDateColumn(row) {
    var i;
    for (i = 0; i < row.length; i++) {
      if (normalizeHeader(row[i]).indexOf('日期') !== -1) return i;
    }
    for (i = 0; i < row.length; i++) {
      if (normalizeHeader(row[i]) === '日') return i;
    }
    return -1;
  }

  function findWeekColumn(row) {
    var keywords = ['星期', '週', '禮拜'];
    for (var i = 0; i < row.length; i++) {
      var cell = normalizeHeader(row[i]);
      for (var k = 0; k < keywords.length; k++) {
        if (cell.indexOf(keywords[k]) !== -1) return i;
      }
    }
    return -1;
  }

  function detectColumns(table, maxScanRows) {
    var limit = Math.min(table.length, maxScanRows || 10);
    for (var r = 0; r < limit; r++) {
      var row = table[r] || [];
      var dateCol = findDateColumn(row);
      var weekCol = findWeekColumn(row);
      if (dateCol === -1 || weekCol === -1 || dateCol === weekCol) continue;

      var checkCols = [];
      for (var c = 0; c < row.length; c++) {
        if (c !== dateCol && c !== weekCol) checkCols.push(c);
      }
      return { headerRow: r, dateCol: dateCol, weekCol: weekCol, checkCols: checkCols };
    }
    return null;
  }

  return {
    SYMBOLS: SYMBOLS,
    isSymbolOnly: isSymbolOnly,
    detectColumns: detectColumns
  };
});
