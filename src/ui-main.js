// 介面主流程：串起選年月、取假日、讀檔、預覽、下載
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else {
    root.UiMain = api;
    if (typeof document !== 'undefined' && document.getElementById('drop')) api.start();
  }
})(typeof self !== 'undefined' ? self : this, function () {

  var ROC_OFFSET = 1911;

  function pad2(n) { return n < 10 ? '0' + n : String(n); }

  function outputFilename(originalName, year, month) {
    var suffix = '_' + (year - ROC_OFFSET) + '年' + pad2(month) + '月';
    var at = originalName.lastIndexOf('.');
    if (at <= 0) return originalName + suffix;
    return originalName.slice(0, at) + suffix + originalName.slice(at);
  }

  function detectFormat(filename) {
    var lower = String(filename).toLowerCase();
    if (/\.xlsx$/.test(lower)) return 'xlsx';
    if (/\.docx$/.test(lower)) return 'docx';
    if (/\.(xls|doc)$/.test(lower)) return 'legacy';
    return 'unknown';
  }

  function storageKey(headers) {
    return 'week-automodify:' + headers.join('|');
  }

  function columnsFrom(headerCells, headerRow, dateCol, weekCol) {
    var checkCols = [];
    for (var c = 0; c < headerCells.length; c++) {
      if (c !== dateCol && c !== weekCol) checkCols.push(c);
    }
    return { headerRow: headerRow, dateCol: dateCol, weekCol: weekCol, checkCols: checkCols };
  }

  function start() {
    var $ = function (id) { return document.getElementById(id); };
    var state = {
      year: null, month: null, model: [], source: '',
      file: null, format: null, parsed: null, cols: null, grid: null,
      plan: null, yearMonthTargets: []
    };

    function esc(v) {
      return XmlText.escape(v === null || v === undefined ? '' : String(v));
    }

    function showError(msg) {
      $('messages').innerHTML = '<div class="error">' + esc(msg) + '</div>';
      $('previewPanel').classList.remove('hidden');
    }

    function initSelects() {
      var d = UiCalendar.defaultYearMonth();
      state.year = d.year;
      state.month = d.month;

      var ys = $('year');
      for (var y = d.year - 2; y <= d.year + 2; y++) {
        var o = document.createElement('option');
        o.value = y;
        o.textContent = y + '年（民國 ' + (y - ROC_OFFSET) + ' 年）';
        if (y === d.year) o.selected = true;
        ys.appendChild(o);
      }
      var ms = $('month');
      for (var m = 1; m <= 12; m++) {
        var o2 = document.createElement('option');
        o2.value = m;
        o2.textContent = m + ' 月';
        if (m === d.month) o2.selected = true;
        ms.appendChild(o2);
      }
      ys.addEventListener('change', function () { state.year = parseInt(ys.value, 10); refreshMonth(); });
      ms.addEventListener('change', function () { state.month = parseInt(ms.value, 10); refreshMonth(); });
    }

    function refreshMonth() {
      $('source').textContent = '讀取假日資料中……';
      HolidaysTW.getYearData(state.year, { builtin: HolidaysBuiltin }).then(function (r) {
        state.source = r.source;
        state.model = MonthModel.buildMonth(state.year, state.month, r.data);
        $('source').textContent = UiCalendar.sourceLabel(r.source, state.year);
        UiCalendar.renderCalendar(state.model, $('calendar'), toggleHoliday);
        if (state.parsed) refreshPreview();
      });
    }

    function toggleHoliday(day) {
      day.isHoliday = !day.isHoliday;
      day.isMakeup = false;
      UiCalendar.renderCalendar(state.model, $('calendar'), toggleHoliday);
      if (state.parsed) refreshPreview();
    }

    function readFile(file) {
      state.file = file;
      state.format = detectFormat(file.name);
      $('fileInfo').textContent = file.name;

      if (state.format === 'legacy') {
        return showError('這是舊版的 Office 格式。請用 Excel 或 Word 開啟後，' +
          '另存成 .xlsx 或 .docx 再試一次。');
      }
      if (state.format === 'unknown') {
        return showError('只支援 .xlsx 與 .docx 兩種格式。');
      }

      file.arrayBuffer().then(function (buf) {
        var loader = state.format === 'xlsx'
          ? ExcelAdapter.loadWorkbook(buf, JSZip)
          : WordAdapter.loadDocument(buf, JSZip);
        return loader;
      }).then(function (parsed) {
        state.parsed = parsed;
        var grid = state.format === 'xlsx' ? parsed.table : (parsed.grids[0] || []);
        if (!grid || grid.length === 0) {
          return showError('這個檔案裡找不到表格。');
        }
        state.grid = grid;
        state.cols = TableAnalysis.detectColumns(grid) || restoreColumns(grid);
        if (!state.cols) {
          return showManualSelect(grid);
        }
        $('manualPanel').classList.add('hidden');
        refreshPreview();
      }).catch(function (err) {
        showError('讀取失敗：' + err.message);
      });
    }

    function readStorage(key) {
      try {
        var raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : null;
      } catch (e) {
        return null;
      }
    }

    function writeStorage(key, value) {
      try {
        localStorage.setItem(key, JSON.stringify(value));
      } catch (e) {
        // 忽略
      }
    }

    function restoreColumns(grid) {
      var limit = Math.min(grid.length, 10);
      for (var r = 0; r < limit; r++) {
        var cells = (grid[r] || []).map(function (x) { return String(x === null || x === undefined ? '' : x); });
        if (cells.length === 0) continue;
        var saved = readStorage(storageKey(cells));
        if (saved) return columnsFrom(cells, r, saved.dateCol, saved.weekCol);
      }
      return null;
    }

    function showManualSelect(grid) {
      var limit = Math.min(grid.length, 10);
      var html = ['<table class="preview"><tbody>'];
      for (var r = 0; r < limit; r++) {
        html.push('<tr><th>第 ' + (r + 1) + ' 列</th>');
        (grid[r] || []).forEach(function (c) { html.push('<td>' + esc(c) + '</td>'); });
        html.push('</tr>');
      }
      html.push('</tbody></table>');
      $('manualTable').innerHTML = html.join('');

      var headerSelect = $('manualHeader');
      headerSelect.innerHTML = '';
      for (var r2 = 0; r2 < limit; r2++) {
        var o = document.createElement('option');
        o.value = r2;
        o.textContent = '第 ' + (r2 + 1) + ' 列';
        headerSelect.appendChild(o);
      }
      fillColSelects();
      headerSelect.onchange = fillColSelects;
      $('manualPanel').classList.remove('hidden');
      $('previewPanel').classList.add('hidden');
    }

    function fillColSelects() {
      var r = parseInt($('manualHeader').value, 10);
      var cells = state.grid[r] || [];
      [$('manualDate'), $('manualWeek')].forEach(function (sel) {
        sel.innerHTML = '';
        cells.forEach(function (c, i) {
          var o = document.createElement('option');
          o.value = i;
          o.textContent = '第 ' + (i + 1) + ' 欄' + (c ? '（' + c + '）' : '');
          sel.appendChild(o);
        });
      });
      if (cells.length > 1) $('manualWeek').value = 1;
    }

    function applyManualSelect() {
      var r = parseInt($('manualHeader').value, 10);
      var dateCol = parseInt($('manualDate').value, 10);
      var weekCol = parseInt($('manualWeek').value, 10);
      if (dateCol === weekCol) {
        return showError('日期欄與星期欄不能是同一欄。');
      }
      var cells = (state.grid[r] || []).map(function (x) { return String(x === null || x === undefined ? '' : x); });
      state.cols = columnsFrom(cells, r, dateCol, weekCol);
      writeStorage(storageKey(cells), { dateCol: dateCol, weekCol: weekCol });
      $('manualPanel').classList.add('hidden');
      refreshPreview();
    }

    function getTable() {
      return state.format === 'xlsx' ? state.parsed.table : state.parsed.grids[0];
    }

    function refreshPreview() {
      var grid = getTable();
      state.plan = WritePlan.buildWritePlan({
        monthModel: state.model, table: grid, headerRow: state.cols.headerRow,
        dateCol: state.cols.dateCol, weekCol: state.cols.weekCol, checkCols: state.cols.checkCols
      });

      state.yearMonthTargets = state.format === 'xlsx'
        ? ExcelAdapter.findYearMonthTargets(state.parsed.sharedStrings)
        : WordAdapter.findParagraphYearMonth(state.parsed.documentXml);

      renderMessages();
      renderYearMonthList();
      renderPreviewTable(grid);
      $('outputName').textContent = '輸出為：' +
        outputFilename(state.file.name, state.year, state.month);
      $('previewPanel').classList.remove('hidden');
    }

    function renderMessages() {
      var parts = [];
      if (state.plan.missingDays.length > 0) {
        parts.push('<div class="warn"><strong>表格列數不足。</strong>這幾天寫不進去：' +
          state.plan.missingDays.join('、') + ' 日。仍可繼續，但這些日子不會出現在檔案裡。</div>');
      }
      if (state.plan.contentLoss.length > 0) {
        var items = state.plan.contentLoss.map(function (c) {
          return '第 ' + (c.row + 1) + ' 列：' + esc(c.text);
        }).join('<br>');
        parts.push('<div class="warn"><strong>下列內容將被清空。</strong><br>' + items + '</div>');
      }
      if (state.yearMonthTargets.length === 0) {
        parts.push('<div class="muted">未在文件中偵測到年月文字，將只填寫表格。</div>');
      }
      $('messages').innerHTML = parts.join('');
    }

    function renderYearMonthList() {
      var box = $('yearMonthList');
      if (state.yearMonthTargets.length === 0) { box.innerHTML = ''; return; }

      var rows = [];
      state.yearMonthTargets.forEach(function (t) {
        var idx = t.index !== undefined ? t.index : t.paraIndex;
        t.matches.forEach(function (m, i) {
          var key = idx + '-' + i;
          var newVal = YearMonthText.rewriteMatch(m, state.year, state.month);
          rows.push('<label style="display:block"><input type="checkbox" class="ym" value="' +
            esc(key) + '" checked> ' + esc(m.raw) + ' → <strong>' + esc(newVal) + '</strong></label>');
        });
      });
      box.innerHTML = '<div class="warn"><strong>文件中的年月將這樣改：</strong>' +
        rows.join('') + '</div>';
    }

    function renderPreviewTable(grid) {
      var headerCells = grid[state.cols.headerRow] || [];
      var html = ['<table class="preview"><thead><tr><th>列</th>'];
      headerCells.forEach(function (h) { html.push('<th>' + esc(h) + '</th>'); });
      html.push('</tr></thead><tbody>');

      var byCell = {};
      state.plan.writes.forEach(function (w) { byCell[w.row + ',' + w.col] = w.value; });

      state.plan.rowSummaries.forEach(function (s) {
        var cls = s.kind === 'holiday' ? 'is-holiday' : (s.kind === 'makeup' ? 'is-makeup' : '');
        html.push('<tr class="' + cls + '"><td>' + (s.row + 1) + '</td>');
        headerCells.forEach(function (_, c) {
          var v = byCell[s.row + ',' + c];
          html.push('<td>' + esc(v) + '</td>');
        });
        html.push('</tr>');
      });
      html.push('</tbody></table>');
      $('preview').innerHTML = html.join('');
    }

    function collectSelectedKeys() {
      var boxes = document.querySelectorAll('input.ym');
      var keys = [];
      for (var i = 0; i < boxes.length; i++) {
        if (boxes[i].checked) keys.push(boxes[i].value);
      }
      return keys;
    }

    function download() {
      var keys = collectSelectedKeys();

      function buildExcel() {
        var applied = ExcelAdapter.applyWritesToSheet(
          state.parsed.sheetXml, state.plan.writes, state.parsed.cells);
        var ss = ExcelAdapter.rewriteSharedStrings(
          state.parsed.sharedStringsXml, state.parsed.sharedStrings,
          state.yearMonthTargets, state.year, state.month, keys);
        return ExcelAdapter.saveWorkbook(
          state.parsed.zip, state.parsed.sheetPath, applied.xml, ss, JSZip, 'blob');
      }

      function buildWord() {
        var xml = WordAdapter.applyWritesToDocument(
          state.parsed.documentXml, state.parsed.tables[0], state.plan.writes);
        var targets = WordAdapter.findParagraphYearMonth(xml);
        xml = WordAdapter.rewriteParagraphYearMonth(xml, targets, state.year, state.month, keys);
        return WordAdapter.saveDocument(state.parsed.zip, xml, JSZip, 'blob');
      }

      var generate = state.format === 'xlsx' ? buildExcel() : buildWord();
      generate.then(function (blob) {
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = outputFilename(state.file.name, state.year, state.month);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
      }).catch(function (err) {
        showError('產生檔案失敗：' + err.message);
      });
    }

    function bindDragDrop() {
      var drop = $('drop');
      var input = $('file');
      drop.addEventListener('click', function () { input.click(); });
      input.addEventListener('change', function () {
        if (input.files[0]) readFile(input.files[0]);
      });
      ['dragenter', 'dragover'].forEach(function (e) {
        drop.addEventListener(e, function (ev) {
          ev.preventDefault(); drop.classList.add('over');
        });
      });
      ['dragleave', 'drop'].forEach(function (e) {
        drop.addEventListener(e, function (ev) {
          ev.preventDefault(); drop.classList.remove('over');
        });
      });
      drop.addEventListener('drop', function (ev) {
        if (ev.dataTransfer.files[0]) readFile(ev.dataTransfer.files[0]);
      });
    }

    initSelects();
    bindDragDrop();
    $('download').addEventListener('click', download);
    $('manualConfirm').addEventListener('click', applyManualSelect);
    refreshMonth();
  }

  return {
    outputFilename: outputFilename,
    detectFormat: detectFormat,
    storageKey: storageKey,
    columnsFrom: columnsFrom,
    start: start
  };
});
