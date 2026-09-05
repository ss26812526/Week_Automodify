// Word 轉接層：以正規式直接處理 docx 內部 XML
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'), require('./year-month-text.js'));
  } else {
    root.WordAdapter = factory(root.XmlText, root.YearMonthText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText, YearMonthText) {

  function textRegex() {
    return /(<w:t\b[^>]*>)([\s\S]*?)<\/w:t>/g;
  }

  function splitSegments(container) {
    var re = textRegex();
    var out = [];
    var m;
    var cursor = 0;
    while ((m = re.exec(container)) !== null) {
      var text = XmlText.unescape(m[2]);
      out.push({
        raw: m[0],
        openTag: m[1],
        text: text,
        start: cursor,
        end: cursor + text.length,
        xmlIndex: m.index
      });
      cursor += text.length;
    }
    return out;
  }

  function mergedText(container) {
    return splitSegments(container).map(function (s) { return s.text; }).join('');
  }

  function parseTables(documentXml) {
    var tables = [];
    var tblRe = /<w:tbl>([\s\S]*?)<\/w:tbl>/g;
    var tm;
    while ((tm = tblRe.exec(documentXml)) !== null) {
      var tblInner = tm[1];
      var rows = [];

      var trRe = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
      var rm;
      while ((rm = trRe.exec(tblInner)) !== null) {
        var trInner = rm[1];
        var cells = [];

        var tcRe = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
        var cm;
        while ((cm = tcRe.exec(trInner)) !== null) {
          cells.push({ raw: cm[0], inner: cm[1], text: mergedText(cm[1]) });
        }
        rows.push({ raw: rm[0], inner: trInner, cells: cells });
      }
      tables.push({ raw: tm[0], inner: tblInner, rows: rows });
    }
    return tables;
  }

  function tableToGrid(table) {
    return table.rows.map(function (row) {
      return row.cells.map(function (c) { return c.text; });
    });
  }

  function loadDocument(data, JSZipLib) {
    return JSZipLib.loadAsync(data).then(function (zip) {
      var file = zip.file('word/document.xml');
      if (!file) throw new Error('這個檔案裡找不到 Word 內文，可能不是有效的 Word 檔');
      return file.async('string').then(function (documentXml) {
        var tables = parseTables(documentXml);
        return {
          zip: zip,
          documentXml: documentXml,
          tables: tables,
          grids: tables.map(tableToGrid)
        };
      });
    });
  }

  function replaceRange(container, segments, from, to, replacement) {
    var touched = segments.filter(function (s) {
      return s.start < to && s.end > from;
    });
    if (touched.length === 0) return container;

    var updates = [];
    touched.forEach(function (seg, i) {
      var before = seg.text.slice(0, Math.max(0, from - seg.start));
      var after = seg.text.slice(Math.max(0, to - seg.start));
      var newText = i === 0 ? before + replacement + after : before + after;
      updates.push({
        raw: seg.raw,
        replacement: seg.openTag + XmlText.escape(newText) + '</w:t>'
      });
    });

    var out = container;
    for (var i = updates.length - 1; i >= 0; i--) {
      var at = out.lastIndexOf(updates[i].raw);
      if (at === -1) continue;
      out = out.slice(0, at) + updates[i].replacement + out.slice(at + updates[i].raw.length);
    }
    return out;
  }

  function setCellText(cellXml, value) {
    var segments = splitSegments(cellXml);

    if (segments.length > 0) {
      var total = segments[segments.length - 1].end;
      if (total > 0) {
        return replaceRange(cellXml, segments, 0, total, value);
      }
      // 所有片段都是空文字，直接把新值寫進第一個片段
      var seg = segments[0];
      var newTag = seg.openTag + XmlText.escape(value) + '</w:t>';
      var at = cellXml.indexOf(seg.raw);
      if (at !== -1) {
        return cellXml.slice(0, at) + newTag + cellXml.slice(at + seg.raw.length);
      }
    }

    if (value === '') return cellXml;
    var newRun = '<w:r><w:t xml:space="preserve">' + XmlText.escape(value) + '</w:t></w:r>';
    var at = cellXml.lastIndexOf('</w:p>');
    if (at === -1) return cellXml;
    return cellXml.slice(0, at) + newRun + cellXml.slice(at);
  }

  function applyWritesToDocument(documentXml, table, writes) {
    var byRow = {};
    writes.forEach(function (w) {
      if (!byRow[w.row]) byRow[w.row] = [];
      byRow[w.row].push(w);
    });

    var newTableXml = table.raw;

    Object.keys(byRow).forEach(function (rowKey) {
      var rowIndex = parseInt(rowKey, 10);
      var row = table.rows[rowIndex];
      if (!row) return;

      var newRowXml = row.raw;
      byRow[rowKey].forEach(function (w) {
        var cell = row.cells[w.col];
        if (!cell) return;
        var newCellXml = setCellText(cell.raw, w.value);
        var at = newRowXml.indexOf(cell.raw);
        if (at === -1) return;
        newRowXml = newRowXml.slice(0, at) + newCellXml + newRowXml.slice(at + cell.raw.length);
      });

      var at = newTableXml.indexOf(row.raw);
      if (at === -1) return;
      newTableXml = newTableXml.slice(0, at) + newRowXml + newTableXml.slice(at + row.raw.length);
    });

    var at = documentXml.indexOf(table.raw);
    if (at === -1) return documentXml;
    return documentXml.slice(0, at) + newTableXml + documentXml.slice(at + table.raw.length);
  }

  function findParagraphYearMonth(documentXml) {
    var out = [];
    var paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
    var m;
    var index = 0;
    while ((m = paraRe.exec(documentXml)) !== null) {
      var raw = m[0];
      var text = mergedText(raw);
      var matches = YearMonthText.findYearMonth(text);
      if (matches.length > 0) {
        out.push({ paraIndex: index, raw: raw, text: text, matches: matches });
      }
      index++;
    }
    return out;
  }

  function rewriteParagraphYearMonth(documentXml, targets, targetYear, targetMonth, selectedKeys) {
    var out = documentXml;

    targets.forEach(function (target) {
      var newParaXml = target.raw;
      var ordered = target.matches
        .map(function (m, i) { return { match: m, i: i }; })
        .sort(function (a, b) { return b.match.index - a.match.index; });

      var changed = false;
      ordered.forEach(function (entry) {
        var key = target.paraIndex + '-' + entry.i;
        if (selectedKeys && selectedKeys.indexOf(key) === -1) return;
        var newText = YearMonthText.rewriteMatch(entry.match, targetYear, targetMonth);
        if (newText === entry.match.raw) return;
        newParaXml = replaceRange(
          newParaXml, splitSegments(newParaXml),
          entry.match.index, entry.match.index + entry.match.raw.length,
          newText
        );
        changed = true;
      });

      if (!changed) return;
      var at = out.indexOf(target.raw);
      if (at === -1) return;
      out = out.slice(0, at) + newParaXml + out.slice(at + target.raw.length);
    });

    return out;
  }

  function saveDocument(zip, documentXml, JSZipLib, outputType) {
    zip.file('word/document.xml', documentXml);
    return zip.generateAsync({
      type: outputType || 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  return {
    splitSegments: splitSegments,
    mergedText: mergedText,
    parseTables: parseTables,
    tableToGrid: tableToGrid,
    loadDocument: loadDocument,
    replaceRange: replaceRange,
    setCellText: setCellText,
    applyWritesToDocument: applyWritesToDocument,
    findParagraphYearMonth: findParagraphYearMonth,
    rewriteParagraphYearMonth: rewriteParagraphYearMonth,
    saveDocument: saveDocument
  };
});
