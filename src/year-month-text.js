// 年月文字的偵測與改寫：辨識「115年8月」「2026年08月」等寫法並沿用原格式改寫
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.YearMonthText = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var ROC_MAX = 200;
  var ROC_OFFSET = 1911;

  function buildPattern() {
    return /([0-9]{2,4})([ ]*)年([ ]*)([0-9]{1,2})([ ]*)月/g;
  }

  function findYearMonth(text) {
    var pattern = buildPattern();
    var out = [];
    var m;
    while ((m = pattern.exec(text)) !== null) {
      var yearRaw = m[1];
      var monthRaw = m[4];
      var yearNum = parseInt(yearRaw, 10);
      var era = yearNum <= ROC_MAX ? 'roc' : 'ad';
      out.push({
        index: m.index,
        raw: m[0],
        yearRaw: yearRaw,
        monthRaw: monthRaw,
        era: era,
        year: era === 'roc' ? yearNum + ROC_OFFSET : yearNum,
        month: parseInt(monthRaw, 10),
        gapAfterYearNum: m[2],
        gapAfterYearChar: m[3],
        gapAfterMonthNum: m[5]
      });
    }
    return out;
  }

  function padLike(sample, value) {
    var s = String(value);
    if (sample.length > s.length && sample.charAt(0) === '0') {
      return '0'.repeat(sample.length - s.length) + s;
    }
    return s;
  }

  function rewriteMatch(match, targetYear, targetMonth) {
    var yearValue = match.era === 'roc' ? targetYear - ROC_OFFSET : targetYear;
    return padLike(match.yearRaw, yearValue) +
      match.gapAfterYearNum + '年' + match.gapAfterYearChar +
      padLike(match.monthRaw, targetMonth) +
      match.gapAfterMonthNum + '月';
  }

  function applyRewrites(text, matches, targetYear, targetMonth, selected) {
    var chosen = matches.filter(function (m, i) {
      return !selected || selected.indexOf(i) !== -1;
    });
    var ordered = chosen.slice().sort(function (a, b) { return b.index - a.index; });
    var out = text;
    ordered.forEach(function (m) {
      out = out.slice(0, m.index) + rewriteMatch(m, targetYear, targetMonth) + out.slice(m.index + m.raw.length);
    });
    return out;
  }

  return {
    ROC_MAX: ROC_MAX,
    ROC_OFFSET: ROC_OFFSET,
    findYearMonth: findYearMonth,
    rewriteMatch: rewriteMatch,
    applyRewrites: applyRewrites
  };
});
