// 開發端工具：產生仿真的機房月檢查表，供測試使用
// 用法：node tools/make-samples.js
var fs = require('fs');
var path = require('path');
var ExcelJS = require('exceljs');
var docx = require('docx');
var Document = docx.Document, Packer = docx.Packer, Paragraph = docx.Paragraph;
var Table = docx.Table, TableRow = docx.TableRow, TableCell = docx.TableCell, TextRun = docx.TextRun;

var headers = ['日期', '星期', '溫濕度', 'UPS', '消防設備', '門禁', '備註'];
var title = '機房狀態月檢查表';
var yearMonth = '115年8月';
var totalRows = 31;
var outDir = path.join(__dirname, '..', 'samples');

function dataRow(day) {
  var remark = day === 3 ? '空調異音已通報維護廠商' : '';
  return [String(day), '', 'V', 'V', 'V', 'V', remark];
}

async function generateExcel() {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('8月');

  ws.mergeCells(1, 1, 1, headers.length);
  ws.getCell(1, 1).value = title + '  ' + yearMonth;
  ws.getCell(1, 1).alignment = { horizontal: 'center' };
  ws.getCell(1, 1).font = { size: 14, bold: true };

  ws.addRow([]);
  ws.addRow(headers);
  for (var d = 1; d <= totalRows; d++) ws.addRow(dataRow(d));

  var thinBorder = { style: 'thin' };
  for (var r = 3; r <= 3 + totalRows; r++) {
    for (var c = 1; c <= headers.length; c++) {
      ws.getCell(r, c).border = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };
    }
  }
  ws.getRow(3).font = { bold: true };

  var p = path.join(outDir, '機房月檢查表-範例.xlsx');
  await wb.xlsx.writeFile(p);
  console.log('已產生 ' + p);
}

function tableRow(cells, bold) {
  return new TableRow({
    children: cells.map(function (t) {
      return new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: t, bold: !!bold })] })]
      });
    })
  });
}

async function generateWord() {
  var rows = [tableRow(headers, true)];
  for (var d = 1; d <= totalRows; d++) rows.push(tableRow(dataRow(d), false));

  var doc = new Document({
    sections: [{
      children: [
        new Paragraph({
          children: [
            new TextRun(title + '  '),
            new TextRun({ text: yearMonth, bold: true })
          ]
        }),
        new Paragraph(''),
        new Table({ rows: rows })
      ]
    }]
  });

  var buf = await Packer.toBuffer(doc);
  var p = path.join(outDir, '機房月檢查表-範例.docx');
  fs.writeFileSync(p, buf);
  console.log('已產生 ' + p);
}

(async function () {
  fs.mkdirSync(outDir, { recursive: true });
  await generateExcel();
  await generateWord();
})();
