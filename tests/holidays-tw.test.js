(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var H = isNode ? require('../src/holidays-tw.js') : self.HolidaysTW;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual;

  var 樣本 = [
    { date: '20250208', week: '六', isHoliday: false, description: '補行上班' },
    { date: '20250209', week: '日', isHoliday: true, description: '' },
    { date: '20250210', week: '一', isHoliday: false, description: '' },
    { date: '20250228', week: '五', isHoliday: true, description: '和平紀念日' }
  ];

  describe('假日資料 › 正規化', function () {
    it('轉成以日期為鍵的物件', function () {
      var d = H.normalizeYearData(樣本);
      assertEqual(d['20250228'].description, '和平紀念日');
      assertEqual(d['20250228'].isHoliday, true);
      assertEqual(d['20250228'].week, '五');
    });

    it('保留補班日為非假日', function () {
      var d = H.normalizeYearData(樣本);
      assertEqual(d['20250208'].isHoliday, false);
      assertEqual(d['20250208'].week, '六');
    });

    it('空陣列產生空物件', function () {
      assertEqual(Object.keys(H.normalizeYearData([])).length, 0);
    });
  });

  describe('假日資料 › 純計算後備', function () {
    it('週六判為假日', function () {
      var d = H.computeYearData(2026);
      assertEqual(d['20260905'].isHoliday, true);
      assertEqual(d['20260905'].week, '六');
    });

    it('週日判為假日', function () {
      assertEqual(H.computeYearData(2026)['20260906'].isHoliday, true);
    });

    it('平日判為非假日', function () {
      assertEqual(H.computeYearData(2026)['20260907'].isHoliday, false);
    });

    it('國定假日無法判斷，中秋節仍視為平日', function () {
      assertEqual(H.computeYearData(2026)['20260925'].isHoliday, false);
    });

    it('平年產生 365 筆，閏年產生 366 筆', function () {
      assertEqual(Object.keys(H.computeYearData(2026)).length, 365);
      assertEqual(Object.keys(H.computeYearData(2024)).length, 366);
    });
  });

  describe('假日資料 › 網址', function () {
    it('依年份組出正確網址', function () {
      assertEqual(H.buildUrl(2026),
        'https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/2026.json');
    });
  });

  describe('假日資料 › 三層遞降', function () {
    it('線上取得成功時來源為 online', async function () {
      var fake = function () {
        return Promise.resolve({ ok: true, json: function () { return Promise.resolve(樣本); } });
      };
      var r = await H.getYearData(2025, { fetchImpl: fake });
      assertEqual(r.source, 'online');
      assertEqual(r.data['20250228'].description, '和平紀念日');
    });

    it('線上失敗時退回內建資料', async function () {
      var fake = function () { return Promise.reject(new Error('網路不通')); };
      var r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
      assertEqual(r.data['20250228'].description, '和平紀念日');
    });

    it('回應狀態非 ok 時也退回內建資料', async function () {
      var fake = function () { return Promise.resolve({ ok: false, status: 404 }); };
      var r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
    });

    it('線上失敗且無內建資料時退回純計算', async function () {
      var fake = function () { return Promise.reject(new Error('網路不通')); };
      var r = await H.getYearData(2026, { fetchImpl: fake, builtin: {} });
      assertEqual(r.source, 'computed');
      assertEqual(r.data['20260905'].isHoliday, true);
    });

    it('逾時視同失敗並退回內建資料', async function () {
      var fake = function () { return new Promise(function () {}); };
      var r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 }, timeoutMs: 30 });
      assertEqual(r.source, 'builtin');
    });

    it('沒有可用的 fetch 實作時直接退回內建資料', async function () {
      var r = await H.getYearData(2025, { fetchImpl: null, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
    });
  });
})();
