// 開發端工具：抓取台灣行事曆並產生 src/holidays-builtin.js
// 用法：node tools/fetch-holidays.js
var fs = require('fs');
var path = require('path');
var H = require('../src/holidays-tw.js');

var YEARS = [2024, 2025, 2026, 2027];

(async function () {
  var collected = {};
  for (var i = 0; i < YEARS.length; i++) {
    var year = YEARS[i];
    var res = await fetch(H.buildUrl(year));
    if (!res.ok) throw new Error(year + ' 年資料取得失敗，狀態 ' + res.status);
    var rows = await res.json();
    collected[year] = rows.map(function (r) {
      return { date: r.date, week: r.week, isHoliday: r.isHoliday, description: r.description || '' };
    });
    console.log(year + ' 年：' + rows.length + ' 筆');
  }

  var today = new Date().toISOString().slice(0, 10);
  var body =
    '// 內建台灣行事曆資料，由 tools/fetch-holidays.js 產生，請勿手動編輯\n' +
    '// 資料來源：https://github.com/ruyut/TaiwanCalendar\n' +
    '// 產生日期：' + today + '\n' +
    '(function (root, factory) {\n' +
    '  if (typeof module === \'object\' && module.exports) module.exports = factory();\n' +
    '  else root.HolidaysBuiltin = factory();\n' +
    '})(typeof self !== \'undefined\' ? self : this, function () {\n' +
    '  return ' + JSON.stringify(collected) + ';\n' +
    '});\n';

  var out = path.join(__dirname, '..', 'src', 'holidays-builtin.js');
  fs.writeFileSync(out, body, 'utf8');
  console.log('已寫入 ' + out + '，共 ' + Math.round(body.length / 1024) + ' KB');
})();
