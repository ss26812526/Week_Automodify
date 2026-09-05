// 極簡測試框架：同時支援 Node 與瀏覽器
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TestRunner = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var tests = [];
  var currentSuite = '';

  function describe(name, fn) {
    var previous = currentSuite;
    currentSuite = previous ? previous + ' › ' + name : name;
    fn();
    currentSuite = previous;
  }

  function it(name, fn) {
    tests.push({ suite: currentSuite, name: name, fn: fn });
  }

  function assertEqual(actual, expected, msg) {
    if (actual !== expected) {
      throw new Error(
        (msg ? msg + '：' : '') +
        '預期 ' + JSON.stringify(expected) + '，實際得到 ' + JSON.stringify(actual)
      );
    }
  }

  function assertDeepEqual(actual, expected, msg) {
    var a = JSON.stringify(actual);
    var e = JSON.stringify(expected);
    if (a !== e) {
      throw new Error((msg ? msg + '：' : '') + '預期 ' + e + '，實際得到 ' + a);
    }
  }

  function assertThrows(fn, msg) {
    var threw = false;
    try { fn(); } catch (err) { threw = true; }
    if (!threw) throw new Error((msg ? msg + '：' : '') + '預期會擲出例外，但沒有');
  }

  async function run() {
    var results = [];
    for (var i = 0; i < tests.length; i++) {
      var t = tests[i];
      try {
        await t.fn();
        results.push({ suite: t.suite, name: t.name, pass: true, error: null });
      } catch (err) {
        results.push({ suite: t.suite, name: t.name, pass: false, error: err.message });
      }
    }
    var passed = results.filter(function (r) { return r.pass; }).length;
    return { total: results.length, passed: passed, failed: results.length - passed, results: results };
  }

  return {
    describe: describe,
    it: it,
    assertEqual: assertEqual,
    assertDeepEqual: assertDeepEqual,
    assertThrows: assertThrows,
    run: run
  };
});
