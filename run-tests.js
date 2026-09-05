// Node 端測試進入點：node run-tests.js
var T = require('./src/test-runner.js');

var testFiles = [
  './tests/test-runner.test.js',
  './tests/calendar-core.test.js',
  './tests/year-month-text.test.js',
  './tests/holidays-tw.test.js',
  './tests/month-model.test.js',
  './tests/table-analysis.test.js',
  './tests/write-plan.test.js',
  './tests/xml-text.test.js',
  './tests/excel-adapter.test.js',
  './tests/word-adapter.test.js',
  './tests/ui-calendar.test.js',
  './tests/ui-main.test.js'
];

testFiles.forEach(function (f) { require(f); });

T.run().then(function (summary) {
  summary.results.forEach(function (r) {
    console.log((r.pass ? '  ✓' : '  ✗') + ' ' + r.suite + ' — ' + r.name);
    if (!r.pass) console.log('      ' + r.error);
  });
  console.log('');
  console.log('共 ' + summary.total + ' 項，通過 ' + summary.passed + ' 項，失敗 ' + summary.failed + ' 項');
  process.exit(summary.failed > 0 ? 1 : 0);
});
