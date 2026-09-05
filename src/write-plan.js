// 寫入計畫：決定每一格該寫什麼，是本工具的業務規則核心
// 不涉及任何檔案格式，Excel 與 Word 共用同一份判斷
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./table-analysis.js'));
  } else {
    root.WritePlan = factory(root.TableAnalysis);
  }
})(typeof self !== 'undefined' ? self : this, function (TableAnalysis) {
  var HOLIDAY_MARK = '-';

  function cellText(table, row, col) {
    var r = table[row];
    if (!r) return '';
    var v = r[col];
    return v === null || v === undefined ? '' : String(v);
  }

  function buildWritePlan(input) {
    var monthModel = input.monthModel;
    var table = input.table;
    var headerRow = input.headerRow;
    var dateCol = input.dateCol;
    var weekCol = input.weekCol;
    var checkCols = input.checkCols || [];

    var writes = [];
    var rowSummaries = [];
    var clearedRows = [];
    var contentLoss = [];

    var firstDataRow = headerRow + 1;
    var dataRowCount = Math.max(0, table.length - firstDataRow);

    function write(row, col, value, checkLoss) {
      if (checkLoss) {
        var original = cellText(table, row, col);
        if (!TableAnalysis.isSymbolOnly(original)) {
          contentLoss.push({ row: row, col: col, text: original });
        }
      }
      writes.push({ row: row, col: col, value: value });
    }

    for (var i = 0; i < dataRowCount; i++) {
      var row = firstDataRow + i;
      var day = monthModel[i];

      if (!day) {
        write(row, dateCol, '', false);
        write(row, weekCol, '', false);
        checkCols.forEach(function (c) { write(row, c, '', true); });
        clearedRows.push(row);
        rowSummaries.push({
          row: row, day: null, weekday: '', isHoliday: false,
          isMakeup: false, description: '', kind: 'cleared'
        });
        continue;
      }

      write(row, dateCol, String(day.day), false);
      write(row, weekCol, day.weekday, false);

      var mark = day.isHoliday ? HOLIDAY_MARK : '';
      checkCols.forEach(function (c) { write(row, c, mark, true); });

      rowSummaries.push({
        row: row,
        day: day.day,
        weekday: day.weekday,
        isHoliday: day.isHoliday,
        isMakeup: day.isMakeup,
        description: day.description,
        kind: day.isHoliday ? 'holiday' : (day.isMakeup ? 'makeup' : 'weekday')
      });
    }

    var missingDays = [];
    for (var j = dataRowCount; j < monthModel.length; j++) {
      missingDays.push(monthModel[j].day);
    }

    return {
      writes: writes,
      rowSummaries: rowSummaries,
      missingDays: missingDays,
      clearedRows: clearedRows,
      contentLoss: contentLoss
    };
  }

  return {
    HOLIDAY_MARK: HOLIDAY_MARK,
    buildWritePlan: buildWritePlan
  };
});
