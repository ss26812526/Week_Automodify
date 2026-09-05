// Excel 轉接層：以正規式直接處理 xlsx 內部 XML，不使用 DOM 解析器
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'), require('./year-month-text.js'));
  } else {
    root.ExcelAdapter = factory(root.XmlText, root.YearMonthText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText, YearMonthText) {

  function colToIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n - 1;
  }

  function indexToCol(index) {
    var n = index + 1;
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function parseCellRef(ref) {
    var m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) return null;
    return { row: parseInt(m[2], 10) - 1, col: colToIndex(m[1]) };
  }

  function parseSharedStrings(xml) {
    var out = [];
    var siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    var m;
    while ((m = siRe.exec(xml)) !== null) {
      var inner = m[1] || '';
      var tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
      var t;
      var text = '';
      while ((t = tRe.exec(inner)) !== null) {
        text += XmlText.unescape(t[1] || '');
      }
      out.push(text);
    }
    return out;
  }

  function cellValue(attrs, inner, sharedStrings) {
    var typeMatch = /\bt="([^"]*)"/.exec(attrs);
    var type = typeMatch ? typeMatch[1] : '';

    if (type === 'inlineStr') {
      var text = '';
      var tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      var t;
      while ((t = tRe.exec(inner)) !== null) text += XmlText.unescape(t[1]);
      return { type: type, text: text };
    }

    var vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
    var rawValue = vMatch ? XmlText.unescape(vMatch[1]) : '';

    if (type === 's') {
      var idx = parseInt(rawValue, 10);
      return { type: type, text: sharedStrings[idx] === undefined ? '' : sharedStrings[idx] };
    }
    return { type: type, text: rawValue };
  }

  function parseSheet(sheetXml, sharedStrings) {
    var table = [];
    var cells = [];
    var rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
    var rowMatch;

    while ((rowMatch = rowRe.exec(sheetXml)) !== null) {
      var rowAttrs = rowMatch[1] !== undefined ? rowMatch[1] : rowMatch[3];
      var rowInner = rowMatch[2] || '';
      var rMatch = /\br="([0-9]+)"/.exec(rowAttrs);
      if (!rMatch) continue;
      var rowIndex = parseInt(rMatch[1], 10) - 1;

      while (table.length <= rowIndex) table.push([]);
      var rowArray = table[rowIndex];

      var cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g;
      var cellMatch;
      while ((cellMatch = cellRe.exec(rowInner)) !== null) {
        var cAttrs = cellMatch[1] !== undefined ? cellMatch[1] : cellMatch[3];
        var cInner = cellMatch[2] || '';
        var refMatch = /\br="([A-Z]+[0-9]+)"/.exec(cAttrs);
        if (!refMatch) continue;
        var ref = refMatch[1];
        var pos = parseCellRef(ref);
        var styleMatch = /\s(s="[0-9]+")/.exec(cAttrs);
        var value = cellValue(cAttrs, cInner, sharedStrings);

        while (rowArray.length <= pos.col) rowArray.push('');
        rowArray[pos.col] = value.text;

        cells.push({
          ref: ref,
          row: pos.row,
          col: pos.col,
          styleAttr: styleMatch ? ' ' + styleMatch[1] : '',
          type: value.type,
          raw: cellMatch[0]
        });
      }
    }
    return { table: table, cells: cells };
  }

  function firstSheetPath(zip) {
    var names = Object.keys(zip.files).filter(function (n) {
      return /^xl\/worksheets\/sheet[0-9]+\.xml$/.test(n);
    });
    names.sort();
    return names[0] || null;
  }

  function loadWorkbook(data, JSZipLib) {
    return JSZipLib.loadAsync(data).then(function (zip) {
      var sheetPath = firstSheetPath(zip);
      if (!sheetPath) throw new Error('這個檔案裡找不到工作表，可能不是有效的 Excel 檔');

      var ssFile = zip.file('xl/sharedStrings.xml');
      var jobs = [zip.file(sheetPath).async('string')];
      jobs.push(ssFile ? ssFile.async('string') : Promise.resolve(''));

      return Promise.all(jobs).then(function (parts) {
        var sheetXml = parts[0];
        var sharedStringsXml = parts[1];
        var sharedStrings = sharedStringsXml ? parseSharedStrings(sharedStringsXml) : [];
        var parsed = parseSheet(sheetXml, sharedStrings);
        return {
          zip: zip,
          sheetPath: sheetPath,
          sheetXml: sheetXml,
          sharedStringsXml: sharedStringsXml,
          sharedStrings: sharedStrings,
          table: parsed.table,
          cells: parsed.cells
        };
      });
    });
  }

  function buildCellXml(ref, styleAttr, value) {
    var style = styleAttr || '';
    if (value === '' || value === null || value === undefined) {
      return '<c r="' + ref + '"' + style + '/>';
    }
    return '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t>' +
      XmlText.escape(value) + '</t></is></c>';
  }

  function applyWritesToSheet(sheetXml, writes, cells) {
    var byPosition = {};
    cells.forEach(function (c) { byPosition[c.row + ',' + c.col] = c; });

    var skipped = [];
    var replacements = [];

    writes.forEach(function (w) {
      var cell = byPosition[w.row + ',' + w.col];
      if (!cell) { skipped.push(w); return; }
      replacements.push({
        from: cell.raw,
        to: buildCellXml(cell.ref, cell.styleAttr, w.value)
      });
    });

    var xml = sheetXml;
    replacements.forEach(function (r) {
      var at = xml.indexOf(r.from);
      if (at === -1) return;
      xml = xml.slice(0, at) + r.to + xml.slice(at + r.from.length);
    });

    return { xml: xml, skipped: skipped };
  }

  function findYearMonthTargets(sharedStrings) {
    var out = [];
    sharedStrings.forEach(function (text, index) {
      var matches = YearMonthText.findYearMonth(text);
      if (matches.length > 0) out.push({ index: index, text: text, matches: matches });
    });
    return out;
  }

  function rewriteSharedStrings(xml, sharedStrings, targets, targetYear, targetMonth, selectedKeys) {
    var out = xml;

    targets.forEach(function (target) {
      var chosen = [];
      target.matches.forEach(function (m, i) {
        var key = target.index + '-' + i;
        if (!selectedKeys || selectedKeys.indexOf(key) !== -1) chosen.push(i);
      });
      if (chosen.length === 0) return;

      var newText = YearMonthText.applyRewrites(
        target.text, target.matches, targetYear, targetMonth, chosen
      );
      if (newText === target.text) return;

      var oldEscaped = XmlText.escape(target.text);
      var newEscaped = XmlText.escape(newText);
      var at = out.indexOf('>' + oldEscaped + '<');
      if (at !== -1) {
        out = out.slice(0, at + 1) + newEscaped + out.slice(at + 1 + oldEscaped.length);
      }
    });

    return out;
  }

  function saveWorkbook(zip, sheetPath, sheetXml, sharedStringsXml, JSZipLib, outputType) {
    zip.file(sheetPath, sheetXml);
    if (sharedStringsXml) zip.file('xl/sharedStrings.xml', sharedStringsXml);
    return zip.generateAsync({
      type: outputType || 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  return {
    colToIndex: colToIndex,
    indexToCol: indexToCol,
    parseCellRef: parseCellRef,
    parseSharedStrings: parseSharedStrings,
    parseSheet: parseSheet,
    loadWorkbook: loadWorkbook,
    buildCellXml: buildCellXml,
    applyWritesToSheet: applyWritesToSheet,
    findYearMonthTargets: findYearMonthTargets,
    rewriteSharedStrings: rewriteSharedStrings,
    saveWorkbook: saveWorkbook
  };
});
