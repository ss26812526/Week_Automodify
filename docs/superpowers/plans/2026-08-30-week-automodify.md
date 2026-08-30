# 機房月檢查表 星期自動填寫工具 實作計畫

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立一個雙擊即開、零安裝、離線可用的單一 HTML 工具，依所選年月自動填寫機房月檢查表的日期與星期，假日列填入 `-`，並同步文件中的年月文字。

**Architecture:** 純前端，無建置工具鏈。底層是零依賴的純函式（日曆、年月文字、寫入計畫），上層是兩個互不相干的檔案轉接層（Excel、Word），皆以 JSZip 解開 Office 的 zip 容器後直接改寫內部 XML，只動儲存格的值，其餘位元組原封不動以完整保留格式。最上層是負責選年月、預覽與下載的介面。

**Tech Stack:** 原生 JavaScript（UMD 模組寫法，同時支援 Node 與瀏覽器）、JSZip 3.10.1（唯一相依套件）、Node.js（僅開發端使用：跑測試、產生範例檔、合成發布檔）

**Spec:** `docs/superpowers/specs/2026-08-30-week-automodify-design.md`

## Global Constraints

- **零相依**：發布檔除 JSZip 外不得引入任何第三方函式庫。**明確禁用 SheetJS**——其社群版寫出 `.xlsx` 不保留儲存格樣式，會使框線與字型遺失。
- **UMD 模組寫法**：`src/` 下每個模組都必須同時可被 Node 的 `require` 與瀏覽器的全域變數取用。測試在兩邊共用。
- **格式保護**：改寫 Office 檔時只更動目標儲存格的值。清空儲存格時必須保留 `<c>` 元素及其 `s` 樣式屬性，僅移除內部的 `<v>` 或 `<is>`。
- **不新增 sharedStrings**：Excel 寫入一律改為行內字串 `t="inlineStr"` 或數字，不得新增 `xl/sharedStrings.xml` 的項目，以免 `count` 與 `uniqueCount` 不符而使檔案損壞。
- **原檔永不覆蓋**：輸出一律使用新檔名，格式為 `原檔名_115年09月.xlsx`。
- **星期字串**：固定使用單字 `一 二 三 四 五 六 日`。
- **星期來源優先序**：優先採用假日資料中的 `week` 欄位，僅在無資料時才由程式推算，以避免時區誤差。
- **假日資料網址**：`https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/<年>.json`
- **註解與介面文字**：一律使用繁體中文。
- **提交訊息**：使用繁體中文，結尾附上 `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`。

---

### Task 1: 專案骨架與測試框架

建立可在 Node 與瀏覽器兩邊執行的極簡測試框架。這是後續每一個任務的驗證基礎，必須先完成。

**Files:**
- Create: `src/test-runner.js`
- Create: `run-tests.js`
- Create: `tests.html`
- Create: `tests/test-runner.test.js`
- Create: `.gitignore`

**Interfaces:**
- Consumes: 無
- Produces:
  - `describe(name: string, fn: () => void): void`
  - `it(name: string, fn: () => void | Promise<void>): void`
  - `assertEqual(actual: any, expected: any, msg?: string): void` — 以 `!==` 比較
  - `assertDeepEqual(actual: any, expected: any, msg?: string): void` — 以 JSON 序列化比較
  - `assertThrows(fn: () => any, msg?: string): void`
  - `run(): Promise<{total, passed, failed, results}>`
  - 瀏覽器全域變數名：`TestRunner`

- [ ] **Step 1: 建立 `.gitignore`**

```
node_modules/
*.log
.DS_Store
Thumbs.db
```

- [ ] **Step 2: 寫測試框架本體 `src/test-runner.js`**

測試採「先收集、後執行」，因此 `it` 的回呼可以是 async 函式。

```javascript
// 極簡測試框架：同時支援 Node 與瀏覽器
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TestRunner = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const tests = [];
  let currentSuite = '';

  function describe(name, fn) {
    const previous = currentSuite;
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
    const a = JSON.stringify(actual);
    const e = JSON.stringify(expected);
    if (a !== e) {
      throw new Error((msg ? msg + '：' : '') + '預期 ' + e + '，實際得到 ' + a);
    }
  }

  function assertThrows(fn, msg) {
    let threw = false;
    try { fn(); } catch (err) { threw = true; }
    if (!threw) throw new Error((msg ? msg + '：' : '') + '預期會擲出例外，但沒有');
  }

  async function run() {
    const results = [];
    for (const t of tests) {
      try {
        await t.fn();
        results.push({ suite: t.suite, name: t.name, pass: true, error: null });
      } catch (err) {
        results.push({ suite: t.suite, name: t.name, pass: false, error: err.message });
      }
    }
    const passed = results.filter(r => r.pass).length;
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
```

- [ ] **Step 3: 寫框架的自我測試 `tests/test-runner.test.js`**

每個測試檔都用這個開頭同時支援兩種環境。後續任務的測試檔一律沿用此樣板。

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual, assertThrows = T.assertThrows;

  describe('測試框架本身', function () {
    it('assertEqual 相等時不擲出例外', function () {
      assertEqual(1, 1);
      assertEqual('一', '一');
    });

    it('assertEqual 不相等時會擲出例外', function () {
      assertThrows(function () { assertEqual(1, 2); });
    });

    it('assertDeepEqual 可比較陣列內容', function () {
      assertDeepEqual([1, 2, 3], [1, 2, 3]);
      assertThrows(function () { assertDeepEqual([1, 2], [1, 3]); });
    });

    it('支援 async 測試', async function () {
      const value = await Promise.resolve(42);
      assertEqual(value, 42);
    });
  });
})();
```

- [ ] **Step 4: 寫 Node 執行進入點 `run-tests.js`**

新增測試檔時必須把檔名加進這個清單。

```javascript
// Node 端測試進入點：node run-tests.js
const T = require('./src/test-runner.js');

const testFiles = [
  './tests/test-runner.test.js'
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
```

- [ ] **Step 5: 寫瀏覽器執行進入點 `tests.html`**

新增測試檔時同樣必須加一行 `<script>`。

```html
<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<title>測試結果</title>
<style>
  body { font-family: "Microsoft JhengHei", sans-serif; margin: 2rem; background: #fafafa; }
  h1 { font-size: 1.2rem; }
  .item { padding: .3rem .6rem; margin: .15rem 0; border-radius: 4px; font-size: .9rem; }
  .pass { background: #e6f7e6; color: #1a5c1a; }
  .fail { background: #fdecea; color: #8c1c13; }
  .err { font-family: monospace; font-size: .8rem; margin-top: .3rem; white-space: pre-wrap; }
  #summary { margin-top: 1.5rem; font-weight: bold; font-size: 1.05rem; }
</style>
</head>
<body>
<h1>測試結果</h1>
<div id="output"></div>
<div id="summary">執行中……</div>

<script src="src/test-runner.js"></script>
<script src="tests/test-runner.test.js"></script>

<script>
TestRunner.run().then(function (summary) {
  const out = document.getElementById('output');
  summary.results.forEach(function (r) {
    const div = document.createElement('div');
    div.className = 'item ' + (r.pass ? 'pass' : 'fail');
    div.textContent = (r.pass ? '✓ ' : '✗ ') + r.suite + ' — ' + r.name;
    if (!r.pass) {
      const e = document.createElement('div');
      e.className = 'err';
      e.textContent = r.error;
      div.appendChild(e);
    }
    out.appendChild(div);
  });
  const s = document.getElementById('summary');
  s.textContent = '共 ' + summary.total + ' 項，通過 ' + summary.passed + ' 項，失敗 ' + summary.failed + ' 項';
  s.style.color = summary.failed > 0 ? '#8c1c13' : '#1a5c1a';
});
</script>
</body>
</html>
```

- [ ] **Step 6: 執行測試確認框架可運作**

Run: `node run-tests.js`
Expected: 4 項全部通過，結尾顯示「共 4 項，通過 4 項，失敗 0 項」，離開碼為 0

- [ ] **Step 7: 提交**

```bash
git add .gitignore src/test-runner.js run-tests.js tests.html tests/test-runner.test.js
git commit -m "建立可在 Node 與瀏覽器共用的極簡測試框架

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 2: 日曆核心

純函式，不碰網路、不碰假日，只負責「這個月有幾天」與「某天是星期幾」。

**Files:**
- Create: `src/calendar-core.js`
- Create: `tests/calendar-core.test.js`
- Modify: `run-tests.js`（testFiles 陣列加入新測試檔）
- Modify: `tests.html`（新增兩行 script 標籤）

**Interfaces:**
- Consumes: 無
- Produces:
  - `WEEKDAY_NAMES: string[]` — `['日','一','二','三','四','五','六']`，索引對應 `Date.getDay()`
  - `daysInMonth(year: number, month: number): number` — month 為 1–12
  - `weekdayOf(year: number, month: number, day: number): string` — 回傳單字如 `'一'`
  - `buildMonthSkeleton(year: number, month: number): Array<{day: number, weekday: string}>`
  - `formatDateKey(year: number, month: number, day: number): string` — 回傳 `'YYYYMMDD'`
  - 瀏覽器全域變數名：`CalendarCore`

- [ ] **Step 1: 寫失敗測試 `tests/calendar-core.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const C = isNode ? require('../src/calendar-core.js') : self.CalendarCore;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('日曆核心 › 月份天數', function () {
    it('平年二月為 28 天', function () {
      assertEqual(C.daysInMonth(2026, 2), 28);
    });

    it('閏年二月為 29 天', function () {
      assertEqual(C.daysInMonth(2024, 2), 29);
    });

    it('世紀年能被 400 整除仍為閏年', function () {
      assertEqual(C.daysInMonth(2000, 2), 29);
    });

    it('世紀年不能被 400 整除則非閏年', function () {
      assertEqual(C.daysInMonth(1900, 2), 28);
    });

    it('小月為 30 天', function () {
      assertEqual(C.daysInMonth(2026, 9), 30);
      assertEqual(C.daysInMonth(2026, 4), 30);
    });

    it('大月為 31 天', function () {
      assertEqual(C.daysInMonth(2026, 1), 31);
      assertEqual(C.daysInMonth(2026, 12), 31);
    });
  });

  describe('日曆核心 › 星期推算', function () {
    it('2026年9月1日為星期二', function () {
      assertEqual(C.weekdayOf(2026, 9, 1), '二');
    });

    it('2026年1月1日為星期四', function () {
      assertEqual(C.weekdayOf(2026, 1, 1), '四');
    });

    it('月初落在星期日', function () {
      assertEqual(C.weekdayOf(2026, 3, 1), '日');
    });

    it('月初落在星期一', function () {
      assertEqual(C.weekdayOf(2026, 6, 1), '一');
    });

    it('跨年後的隔年一月一日星期正確', function () {
      assertEqual(C.weekdayOf(2027, 1, 1), '五');
    });

    it('閏年二月二十九日星期正確', function () {
      assertEqual(C.weekdayOf(2024, 2, 29), '四');
    });
  });

  describe('日曆核心 › 月份骨架', function () {
    it('小月產生 30 筆', function () {
      assertEqual(C.buildMonthSkeleton(2026, 9).length, 30);
    });

    it('平年二月產生 28 筆', function () {
      assertEqual(C.buildMonthSkeleton(2026, 2).length, 28);
    });

    it('第一筆與最後一筆內容正確', function () {
      const rows = C.buildMonthSkeleton(2026, 9);
      assertDeepEqual(rows[0], { day: 1, weekday: '二' });
      assertDeepEqual(rows[29], { day: 30, weekday: '三' });
    });
  });

  describe('日曆核心 › 日期鍵值', function () {
    it('月與日均補零至兩位', function () {
      assertEqual(C.formatDateKey(2026, 9, 5), '20260905');
    });

    it('兩位數的月與日不再補零', function () {
      assertEqual(C.formatDateKey(2026, 12, 25), '20261225');
    });
  });
})();
```

- [ ] **Step 2: 把測試檔接上兩個進入點**

在 `run-tests.js` 的 `testFiles` 陣列加入：

```javascript
  './tests/calendar-core.test.js'
```

在 `tests.html` 的 `src/test-runner.js` 那行之後加入：

```html
<script src="src/calendar-core.js"></script>
<script src="tests/calendar-core.test.js"></script>
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，錯誤訊息為找不到模組 `../src/calendar-core.js`

- [ ] **Step 4: 寫實作 `src/calendar-core.js`**

星期以 `Date.UTC` 推算而非本地時間，避免時區造成跨日誤差。

```javascript
// 日曆核心：零依賴純函式，只處理天數與星期
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CalendarCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  // 索引對應 Date.getUTCDay() 的回傳值，0 為星期日
  const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];

  function isLeapYear(year) {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  }

  function daysInMonth(year, month) {
    const table = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (month === 2 && isLeapYear(year)) return 29;
    return table[month - 1];
  }

  // 使用 UTC 避免本地時區造成跨日誤差
  function weekdayOf(year, month, day) {
    const d = new Date(Date.UTC(year, month - 1, day));
    return WEEKDAY_NAMES[d.getUTCDay()];
  }

  function buildMonthSkeleton(year, month) {
    const total = daysInMonth(year, month);
    const rows = [];
    for (let day = 1; day <= total; day++) {
      rows.push({ day: day, weekday: weekdayOf(year, month, day) });
    }
    return rows;
  }

  function pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  function formatDateKey(year, month, day) {
    return String(year) + pad2(month) + pad2(day);
  }

  return {
    WEEKDAY_NAMES: WEEKDAY_NAMES,
    isLeapYear: isLeapYear,
    daysInMonth: daysInMonth,
    weekdayOf: weekdayOf,
    buildMonthSkeleton: buildMonthSkeleton,
    formatDateKey: formatDateKey
  };
});
```

- [ ] **Step 5: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，共 21 項全數通過

- [ ] **Step 6: 在瀏覽器確認同一批測試也通過**

用瀏覽器開啟 `tests.html`（雙擊即可，不需要 server），確認畫面全綠且顯示「共 21 項，通過 21 項，失敗 0 項」。這一步驗證 UMD 寫法在兩種環境都成立。

- [ ] **Step 7: 提交**

```bash
git add src/calendar-core.js tests/calendar-core.test.js run-tests.js tests.html
git commit -m "新增日曆核心：月份天數與星期推算

以 UTC 推算星期避免時區跨日誤差。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---
### Task 3: 年月文字偵測與改寫

純函式，處理「115年8月」這類文字的辨識與改寫。Excel 與 Word 兩邊共用。

**Files:**
- Create: `src/year-month-text.js`
- Create: `tests/year-month-text.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: 無
- Produces:
  - `findYearMonth(text: string): Array<Match>`，其中 `Match = {index, raw, yearRaw, monthRaw, era, year, month, gapAfterYearNum, gapAfterYearChar, gapAfterMonthNum}`
    - `era` 為 `'roc'`（民國）或 `'ad'`（西元）
    - `year` 一律換算為西元年
  - `rewriteMatch(match: Match, targetYear: number, targetMonth: number): string` — 回傳沿用原格式的新字串
  - `applyRewrites(text: string, matches: Match[], targetYear: number, targetMonth: number, selected?: number[]): string`
    - `selected` 為要套用的 match 索引陣列；省略則全部套用
  - 瀏覽器全域變數名：`YearMonthText`

**紀年判斷的分界**：spec 寫「小於等於 200 視為民國、大於等於 1911 視為西元」，中間區段未定義。本實作採單一分界線：**年份數字小於等於 200 即為民國，否則為西元**，使判斷無空隙。

- [ ] **Step 1: 寫失敗測試 `tests/year-month-text.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const Y = isNode ? require('../src/year-month-text.js') : self.YearMonthText;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('年月文字 › 偵測', function () {
    it('找得到民國年月', function () {
      const m = Y.findYearMonth('機房狀態檢查表 115年8月');
      assertEqual(m.length, 1);
      assertEqual(m[0].era, 'roc');
      assertEqual(m[0].year, 2026);
      assertEqual(m[0].month, 8);
      assertEqual(m[0].raw, '115年8月');
    });

    it('找得到西元年月', function () {
      const m = Y.findYearMonth('2026年08月 檢查紀錄');
      assertEqual(m.length, 1);
      assertEqual(m[0].era, 'ad');
      assertEqual(m[0].year, 2026);
      assertEqual(m[0].month, 8);
    });

    it('一段文字中的多處年月全部找得到', function () {
      const m = Y.findYearMonth('115年8月報表　　製表日期：115年9月');
      assertEqual(m.length, 2);
      assertEqual(m[0].month, 8);
      assertEqual(m[1].month, 9);
    });

    it('沒有年月時回傳空陣列', function () {
      assertDeepEqual(Y.findYearMonth('機房狀態檢查表'), []);
    });

    it('年份 200 視為民國', function () {
      assertEqual(Y.findYearMonth('200年1月')[0].era, 'roc');
    });

    it('年份 201 視為西元', function () {
      assertEqual(Y.findYearMonth('201年1月')[0].era, 'ad');
    });

    it('記錄下原文中的空白', function () {
      const m = Y.findYearMonth('115 年 8 月');
      assertEqual(m[0].raw, '115 年 8 月');
    });
  });

  describe('年月文字 › 改寫', function () {
    it('民國年改寫後仍為民國年', function () {
      const m = Y.findYearMonth('115年8月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '115年9月');
    });

    it('跨年時民國年份跟著進位', function () {
      const m = Y.findYearMonth('114年12月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 1), '115年1月');
    });

    it('西元年改寫後仍為西元年', function () {
      const m = Y.findYearMonth('2025年12月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 1), '2026年1月');
    });

    it('原本月份補零則改寫後也補零', function () {
      const m = Y.findYearMonth('2026年08月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '2026年09月');
    });

    it('原本月份未補零則改寫後也不補零', function () {
      const m = Y.findYearMonth('115年8月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 12), '115年12月');
    });

    it('原文的空白完整沿用', function () {
      const m = Y.findYearMonth('115 年 8 月')[0];
      assertEqual(Y.rewriteMatch(m, 2026, 9), '115 年 9 月');
    });
  });

  describe('年月文字 › 整段套用', function () {
    it('多處年月可一次全部改寫', function () {
      const text = '115年8月報表　　製表：115年8月';
      const m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 9), '115年9月報表　　製表：115年9月');
    });

    it('可只改寫選定的其中一處', function () {
      const text = '115年8月報表　　製表：115年8月';
      const m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 9, [0]), '115年9月報表　　製表：115年8月');
    });

    it('改寫後長度變動不影響其他處的位置', function () {
      const text = '115年9月　115年9月';
      const m = Y.findYearMonth(text);
      assertEqual(Y.applyRewrites(text, m, 2026, 12), '115年12月　115年12月');
    });

    it('沒有匹配時原文不變', function () {
      assertEqual(Y.applyRewrites('機房檢查表', [], 2026, 9), '機房檢查表');
    });
  });
})();
```

- [ ] **Step 2: 把測試檔接上兩個進入點**

`run-tests.js` 的 `testFiles` 加入 `'./tests/year-month-text.test.js'`；`tests.html` 加入：

```html
<script src="src/year-month-text.js"></script>
<script src="tests/year-month-text.test.js"></script>
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/year-month-text.js`

- [ ] **Step 4: 寫實作 `src/year-month-text.js`**

```javascript
// 年月文字的偵測與改寫：辨識「115年8月」「2026年08月」等寫法並沿用原格式改寫
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.YearMonthText = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  // 年份數字小於等於此值視為民國紀年
  const ROC_MAX = 200;
  const ROC_OFFSET = 1911;

  function buildPattern() {
    // 每次都建新的物件，避免 lastIndex 在多次呼叫間殘留
    return /([0-9]{2,4})([ ]*)年([ ]*)([0-9]{1,2})([ ]*)月/g;
  }

  function findYearMonth(text) {
    const pattern = buildPattern();
    const out = [];
    let m;
    while ((m = pattern.exec(text)) !== null) {
      const yearRaw = m[1];
      const monthRaw = m[4];
      const yearNum = parseInt(yearRaw, 10);
      const era = yearNum <= ROC_MAX ? 'roc' : 'ad';
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

  // 若原文有前導零則沿用其位數，否則不補零
  function padLike(sample, value) {
    const s = String(value);
    if (sample.length > s.length && sample.charAt(0) === '0') {
      while (s.length < sample.length) return '0'.repeat(sample.length - s.length) + s;
    }
    return s;
  }

  function rewriteMatch(match, targetYear, targetMonth) {
    const yearValue = match.era === 'roc' ? targetYear - ROC_OFFSET : targetYear;
    return padLike(match.yearRaw, yearValue) +
      match.gapAfterYearNum + '年' + match.gapAfterYearChar +
      padLike(match.monthRaw, targetMonth) +
      match.gapAfterMonthNum + '月';
  }

  function applyRewrites(text, matches, targetYear, targetMonth, selected) {
    const chosen = matches.filter(function (m, i) {
      return !selected || selected.indexOf(i) !== -1;
    });
    // 由後往前替換，避免長度改變影響前面各處的 index
    const ordered = chosen.slice().sort(function (a, b) { return b.index - a.index; });
    let out = text;
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
```

- [ ] **Step 5: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 17 項全數通過

- [ ] **Step 6: 提交**

```bash
git add src/year-month-text.js tests/year-month-text.test.js run-tests.js tests.html
git commit -m "新增年月文字的偵測與改寫

支援民國與西元兩種紀年，改寫時沿用原本的補零與空白格式。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 4: 假日資料層

三層遞降取得台灣行事曆。外部相依（網路、內建資料）一律以參數注入，讓測試完全不碰網路。

**Files:**
- Create: `src/holidays-tw.js`
- Create: `tests/holidays-tw.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: `CalendarCore.daysInMonth`、`CalendarCore.weekdayOf`、`CalendarCore.formatDateKey`
- Produces:
  - `YearData = { [dateKey: string]: {week: string, isHoliday: boolean, description: string} }`
  - `normalizeYearData(raw: Array<{date, week, isHoliday, description}>): YearData`
  - `computeYearData(year: number): YearData` — 只依六日判斷，無國定假日
  - `buildUrl(year: number): string`
  - `fetchOnlineYear(year, fetchImpl, timeoutMs): Promise<YearData>` — 失敗或逾時則 reject
  - `getYearData(year, options): Promise<{source: 'online'|'builtin'|'computed', data: YearData}>`
    - `options = {builtin?: {[year]: Array}, fetchImpl?: Function, timeoutMs?: number}`
  - 瀏覽器全域變數名：`HolidaysTW`

- [ ] **Step 1: 寫失敗測試 `tests/holidays-tw.test.js`**

測試中的假日資料取自實際來源，補班日為 2025-02-08（週六補行上班）。

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const H = isNode ? require('../src/holidays-tw.js') : self.HolidaysTW;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  const 樣本 = [
    { date: '20250208', week: '六', isHoliday: false, description: '補行上班' },
    { date: '20250209', week: '日', isHoliday: true, description: '' },
    { date: '20250210', week: '一', isHoliday: false, description: '' },
    { date: '20250228', week: '五', isHoliday: true, description: '和平紀念日' }
  ];

  describe('假日資料 › 正規化', function () {
    it('轉成以日期為鍵的物件', function () {
      const d = H.normalizeYearData(樣本);
      assertEqual(d['20250228'].description, '和平紀念日');
      assertEqual(d['20250228'].isHoliday, true);
      assertEqual(d['20250228'].week, '五');
    });

    it('保留補班日為非假日', function () {
      const d = H.normalizeYearData(樣本);
      assertEqual(d['20250208'].isHoliday, false);
      assertEqual(d['20250208'].week, '六');
    });

    it('空陣列產生空物件', function () {
      assertEqual(Object.keys(H.normalizeYearData([])).length, 0);
    });
  });

  describe('假日資料 › 純計算後備', function () {
    it('週六判為假日', function () {
      const d = H.computeYearData(2026);
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
      const fake = function () {
        return Promise.resolve({ ok: true, json: function () { return Promise.resolve(樣本); } });
      };
      const r = await H.getYearData(2025, { fetchImpl: fake });
      assertEqual(r.source, 'online');
      assertEqual(r.data['20250228'].description, '和平紀念日');
    });

    it('線上失敗時退回內建資料', async function () {
      const fake = function () { return Promise.reject(new Error('網路不通')); };
      const r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
      assertEqual(r.data['20250228'].description, '和平紀念日');
    });

    it('回應狀態非 ok 時也退回內建資料', async function () {
      const fake = function () { return Promise.resolve({ ok: false, status: 404 }); };
      const r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
    });

    it('線上失敗且無內建資料時退回純計算', async function () {
      const fake = function () { return Promise.reject(new Error('網路不通')); };
      const r = await H.getYearData(2026, { fetchImpl: fake, builtin: {} });
      assertEqual(r.source, 'computed');
      assertEqual(r.data['20260905'].isHoliday, true);
    });

    it('逾時視同失敗並退回內建資料', async function () {
      const fake = function () { return new Promise(function () { /* 永不 resolve */ }); };
      const r = await H.getYearData(2025, { fetchImpl: fake, builtin: { 2025: 樣本 }, timeoutMs: 30 });
      assertEqual(r.source, 'builtin');
    });

    it('沒有可用的 fetch 實作時直接退回內建資料', async function () {
      const r = await H.getYearData(2025, { fetchImpl: null, builtin: { 2025: 樣本 } });
      assertEqual(r.source, 'builtin');
    });
  });
})();
```

- [ ] **Step 2: 把測試檔接上兩個進入點**

`run-tests.js` 加入 `'./tests/holidays-tw.test.js'`；`tests.html` 加入：

```html
<script src="src/holidays-tw.js"></script>
<script src="tests/holidays-tw.test.js"></script>
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/holidays-tw.js`

- [ ] **Step 4: 寫實作 `src/holidays-tw.js`**

```javascript
// 假日資料層：線上取得 → 內建資料 → 純計算，三層遞降
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./calendar-core.js'));
  } else {
    root.HolidaysTW = factory(root.CalendarCore);
  }
})(typeof self !== 'undefined' ? self : this, function (CalendarCore) {
  const BASE_URL = 'https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/';
  const DEFAULT_TIMEOUT_MS = 4000;

  function buildUrl(year) {
    return BASE_URL + year + '.json';
  }

  function normalizeYearData(raw) {
    const out = {};
    (raw || []).forEach(function (row) {
      out[row.date] = {
        week: row.week,
        isHoliday: !!row.isHoliday,
        description: row.description || ''
      };
    });
    return out;
  }

  // 無任何資料來源時的後備：只認得週六與週日
  function computeYearData(year) {
    const out = {};
    for (let month = 1; month <= 12; month++) {
      const total = CalendarCore.daysInMonth(year, month);
      for (let day = 1; day <= total; day++) {
        const week = CalendarCore.weekdayOf(year, month, day);
        out[CalendarCore.formatDateKey(year, month, day)] = {
          week: week,
          isHoliday: week === '六' || week === '日',
          description: ''
        };
      }
    }
    return out;
  }

  function withTimeout(promise, timeoutMs) {
    return new Promise(function (resolve, reject) {
      const timer = setTimeout(function () { reject(new Error('連線逾時')); }, timeoutMs);
      promise.then(
        function (v) { clearTimeout(timer); resolve(v); },
        function (e) { clearTimeout(timer); reject(e); }
      );
    });
  }

  function fetchOnlineYear(year, fetchImpl, timeoutMs) {
    if (!fetchImpl) return Promise.reject(new Error('此環境沒有可用的 fetch'));
    const request = Promise.resolve(fetchImpl(buildUrl(year))).then(function (res) {
      if (!res || !res.ok) throw new Error('回應狀態異常：' + (res ? res.status : '無回應'));
      return res.json();
    }).then(normalizeYearData);
    return withTimeout(request, timeoutMs || DEFAULT_TIMEOUT_MS);
  }

  function getYearData(year, options) {
    const opts = options || {};
    const fetchImpl = opts.fetchImpl !== undefined
      ? opts.fetchImpl
      : (typeof fetch !== 'undefined' ? fetch : null);
    const builtin = opts.builtin || {};

    return fetchOnlineYear(year, fetchImpl, opts.timeoutMs)
      .then(function (data) {
        return { source: 'online', data: data };
      })
      .catch(function () {
        if (builtin[year]) {
          return { source: 'builtin', data: normalizeYearData(builtin[year]) };
        }
        return { source: 'computed', data: computeYearData(year) };
      });
  }

  return {
    BASE_URL: BASE_URL,
    buildUrl: buildUrl,
    normalizeYearData: normalizeYearData,
    computeYearData: computeYearData,
    fetchOnlineYear: fetchOnlineYear,
    getYearData: getYearData
  };
});
```

- [ ] **Step 5: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 15 項全數通過

- [ ] **Step 6: 額外做一次真實網路驗證（不寫成測試）**

這一步確認線上來源仍然可用。因為它依賴外部網路，不可寫進測試套件，否則離線時測試會失敗。

Run:
```bash
node -e "const H=require('./src/holidays-tw.js'); H.getYearData(2026,{}).then(r=>console.log(r.source, r.data['20260925']));"
```
Expected: 印出 `online { week: '五', isHoliday: true, description: '中秋節' }`

- [ ] **Step 7: 提交**

```bash
git add src/holidays-tw.js tests/holidays-tw.test.js run-tests.js tests.html
git commit -m "新增假日資料層：線上、內建、純計算三層遞降

外部相依以參數注入，測試完全不碰網路。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 5: 內建假日資料與月份模型

產生離線用的內建資料，並把日曆骨架與假日資料合併成介面與寫入計畫共用的月份模型。

**Files:**
- Create: `tools/fetch-holidays.js`
- Create: `src/holidays-builtin.js`（由上述腳本產生）
- Create: `src/month-model.js`
- Create: `tests/month-model.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: `CalendarCore.buildMonthSkeleton`、`CalendarCore.formatDateKey`、`HolidaysTW.YearData`
- Produces:
  - `DayModel = {day, weekday, isHoliday, isMakeup, description}`
  - `buildMonth(year: number, month: number, yearData: YearData): DayModel[]`
  - `HolidaysBuiltin: {[year: string]: Array<{date, week, isHoliday, description}>}`（瀏覽器全域變數名 `HolidaysBuiltin`）
  - 瀏覽器全域變數名：`MonthModel`

**`isMakeup` 的定義**：`week` 為六或日、但 `isHoliday` 為 `false`。這種日子要照常檢查，不填 `-`。

- [ ] **Step 1: 寫產生內建資料的腳本 `tools/fetch-holidays.js`**

```javascript
// 開發端工具：抓取台灣行事曆並產生 src/holidays-builtin.js
// 用法：node tools/fetch-holidays.js
const fs = require('fs');
const path = require('path');
const H = require('../src/holidays-tw.js');

const YEARS = [2024, 2025, 2026, 2027];

(async function () {
  const collected = {};
  for (const year of YEARS) {
    const res = await fetch(H.buildUrl(year));
    if (!res.ok) throw new Error(year + ' 年資料取得失敗，狀態 ' + res.status);
    const rows = await res.json();
    // 只留下必要欄位，縮小發布檔體積
    collected[year] = rows.map(function (r) {
      return { date: r.date, week: r.week, isHoliday: r.isHoliday, description: r.description || '' };
    });
    console.log(year + ' 年：' + rows.length + ' 筆');
  }

  const today = new Date().toISOString().slice(0, 10);
  const body =
    '// 內建台灣行事曆資料，由 tools/fetch-holidays.js 產生，請勿手動編輯\n' +
    '// 資料來源：https://github.com/ruyut/TaiwanCalendar\n' +
    '// 產生日期：' + today + '\n' +
    '(function (root, factory) {\n' +
    '  if (typeof module === \'object\' && module.exports) module.exports = factory();\n' +
    '  else root.HolidaysBuiltin = factory();\n' +
    '})(typeof self !== \'undefined\' ? self : this, function () {\n' +
    '  return ' + JSON.stringify(collected) + ';\n' +
    '});\n';

  const out = path.join(__dirname, '..', 'src', 'holidays-builtin.js');
  fs.writeFileSync(out, body, 'utf8');
  console.log('已寫入 ' + out + '，共 ' + Math.round(body.length / 1024) + ' KB');
})();
```

- [ ] **Step 2: 執行腳本產生內建資料**

Run: `node tools/fetch-holidays.js`
Expected: 印出四個年度的筆數（2024 年為 366 筆，其餘為 365 筆），並產生 `src/holidays-builtin.js`

- [ ] **Step 3: 確認產生的檔案可被載入**

Run:
```bash
node -e "const B=require('./src/holidays-builtin.js'); console.log(Object.keys(B), B['2025'].find(x=>x.date==='20250208'));"
```
Expected: 印出 `[ '2024', '2025', '2026', '2027' ]` 以及 2025-02-08 補行上班那筆資料

- [ ] **Step 4: 寫失敗測試 `tests/month-model.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const M = isNode ? require('../src/month-model.js') : self.MonthModel;
  const H = isNode ? require('../src/holidays-tw.js') : self.HolidaysTW;
  const B = isNode ? require('../src/holidays-builtin.js') : self.HolidaysBuiltin;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  const 二〇二六 = H.normalizeYearData(B['2026']);
  const 二〇二五 = H.normalizeYearData(B['2025']);

  describe('月份模型', function () {
    it('小月產生 30 筆', function () {
      assertEqual(M.buildMonth(2026, 9, 二〇二六).length, 30);
    });

    it('平日不是假日也不是補班日', function () {
      const 第一天 = M.buildMonth(2026, 9, 二〇二六)[0];
      assertEqual(第一天.day, 1);
      assertEqual(第一天.weekday, '二');
      assertEqual(第一天.isHoliday, false);
      assertEqual(第一天.isMakeup, false);
    });

    it('週六判為假日', function () {
      const 第五天 = M.buildMonth(2026, 9, 二〇二六)[4];
      assertEqual(第五天.day, 5);
      assertEqual(第五天.weekday, '六');
      assertEqual(第五天.isHoliday, true);
    });

    it('落在平日的國定假日判為假日並帶出名稱', function () {
      const 中秋 = M.buildMonth(2026, 9, 二〇二六)[24];
      assertEqual(中秋.day, 25);
      assertEqual(中秋.weekday, '五');
      assertEqual(中秋.isHoliday, true);
      assertEqual(中秋.description, '中秋節');
    });

    it('教師節判為假日', function () {
      const 教師節 = M.buildMonth(2026, 9, 二〇二六)[27];
      assertEqual(教師節.day, 28);
      assertEqual(教師節.isHoliday, true);
    });

    it('補班日判為非假日且標記為補班', function () {
      const 二月 = M.buildMonth(2025, 2, 二〇二五);
      const 補班 = 二月[7];
      assertEqual(補班.day, 8);
      assertEqual(補班.weekday, '六');
      assertEqual(補班.isHoliday, false);
      assertEqual(補班.isMakeup, true);
    });

    it('缺少該日資料時退回以六日判斷', function () {
      const 空資料 = {};
      const 九月 = M.buildMonth(2026, 9, 空資料);
      assertEqual(九月[4].isHoliday, true, '9月5日為週六');
      assertEqual(九月[0].isHoliday, false, '9月1日為週二');
      assertEqual(九月[24].isHoliday, false, '無資料時中秋節無法辨識');
    });

    it('閏年二月產生 29 筆', function () {
      assertEqual(M.buildMonth(2024, 2, {}).length, 29);
    });
  });
})();
```

- [ ] **Step 5: 把測試檔接上兩個進入點**

`run-tests.js` 加入 `'./tests/month-model.test.js'`；`tests.html` 加入：

```html
<script src="src/holidays-builtin.js"></script>
<script src="src/month-model.js"></script>
<script src="tests/month-model.test.js"></script>
```

- [ ] **Step 6: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/month-model.js`

- [ ] **Step 7: 寫實作 `src/month-model.js`**

```javascript
// 月份模型：把日曆骨架與假日資料合併成介面與寫入計畫共用的結構
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./calendar-core.js'));
  } else {
    root.MonthModel = factory(root.CalendarCore);
  }
})(typeof self !== 'undefined' ? self : this, function (CalendarCore) {

  function buildMonth(year, month, yearData) {
    const data = yearData || {};
    return CalendarCore.buildMonthSkeleton(year, month).map(function (row) {
      const key = CalendarCore.formatDateKey(year, month, row.day);
      const found = data[key];

      if (!found) {
        // 沒有該日資料時只能依六日判斷
        const 是週末 = row.weekday === '六' || row.weekday === '日';
        return {
          day: row.day,
          weekday: row.weekday,
          isHoliday: 是週末,
          isMakeup: false,
          description: ''
        };
      }

      // 星期優先採用資料來源提供的值，避免時區誤差
      const weekday = found.week || row.weekday;
      const 是週末 = weekday === '六' || weekday === '日';
      return {
        day: row.day,
        weekday: weekday,
        isHoliday: found.isHoliday,
        isMakeup: 是週末 && !found.isHoliday,
        description: found.description || ''
      };
    });
  }

  return { buildMonth: buildMonth };
});
```

- [ ] **Step 8: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 8 項全數通過

- [ ] **Step 9: 提交**

```bash
git add tools/fetch-holidays.js src/holidays-builtin.js src/month-model.js tests/month-model.test.js run-tests.js tests.html
git commit -m "新增內建假日資料與月份模型

內建 2024 至 2027 年行事曆供離線使用，
月份模型合併日曆骨架與假日資料並標記補班日。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---
### Task 6: 表格分析

純函式，判斷哪一欄是日期、哪一欄是星期，以及一段文字是否只是勾記符號。Excel 與 Word 共用。

**Files:**
- Create: `src/table-analysis.js`
- Create: `tests/table-analysis.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: 無
- Produces:
  - `SYMBOLS: string[]` — 視為「無資訊」的記號清單
  - `isSymbolOnly(text: string): boolean` — 空字串、空白、或單一記號皆回傳 `true`
  - `detectColumns(table: string[][], maxScanRows?: number): {headerRow, dateCol, weekCol, checkCols} | null`
    - `checkCols` 為該表頭列中除日期欄與星期欄外的所有欄索引
    - 偵測不到時回傳 `null`，由介面轉為手動指定
  - 瀏覽器全域變數名：`TableAnalysis`

**表頭判定規則**：掃描前 `maxScanRows`（預設 10）列，找出同時具備日期欄與星期欄的第一列。日期欄優先比對「日期」，找不到才比對單獨的「日」；星期欄比對「星期」「週」「禮拜」。

- [ ] **Step 1: 寫失敗測試 `tests/table-analysis.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const A = isNode ? require('../src/table-analysis.js') : self.TableAnalysis;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('表格分析 › 記號判定', function () {
    it('空字串視為無資訊', function () {
      assertEqual(A.isSymbolOnly(''), true);
    });

    it('純空白視為無資訊', function () {
      assertEqual(A.isSymbolOnly('   '), true);
    });

    it('勾記符號視為無資訊', function () {
      ['V', 'v', 'X', 'x', '○', '✓', '✗', '-', '－'].forEach(function (s) {
        assertEqual(A.isSymbolOnly(s), true, s + ' 應視為記號');
      });
    });

    it('記號前後有空白仍視為無資訊', function () {
      assertEqual(A.isSymbolOnly(' V '), true);
    });

    it('真正的文字不視為記號', function () {
      assertEqual(A.isSymbolOnly('空調異音已通報'), false);
    });

    it('多個字元的內容不視為記號', function () {
      assertEqual(A.isSymbolOnly('VV'), false);
      assertEqual(A.isSymbolOnly('OK'), false);
    });

    it('null 與 undefined 視為無資訊', function () {
      assertEqual(A.isSymbolOnly(null), true);
      assertEqual(A.isSymbolOnly(undefined), true);
    });
  });

  describe('表格分析 › 欄位偵測', function () {
    const 標準表 = [
      ['機房狀態月檢查表  115年8月', '', '', '', '', ''],
      [],
      ['日期', '星期', '溫濕度', 'UPS', '消防', '備註'],
      ['1', '六', 'V', 'V', 'V', '']
    ];

    it('找得到表頭列與日期星期欄', function () {
      const r = A.detectColumns(標準表);
      assertEqual(r.headerRow, 2);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
    });

    it('其餘欄位皆列為檢查欄', function () {
      assertDeepEqual(A.detectColumns(標準表).checkCols, [2, 3, 4, 5]);
    });

    it('日期欄不在第一欄也找得到', function () {
      const t = [['項次', '日期', '星期', '溫濕度']];
      const r = A.detectColumns(t);
      assertEqual(r.dateCol, 1);
      assertEqual(r.weekCol, 2);
      assertDeepEqual(r.checkCols, [0, 3]);
    });

    it('星期欄寫成「週」也找得到', function () {
      const r = A.detectColumns([['日期', '週', '項目']]);
      assertEqual(r.weekCol, 1);
    });

    it('日期欄只寫「日」也找得到', function () {
      const r = A.detectColumns([['日', '星期', '項目']]);
      assertEqual(r.dateCol, 0);
    });

    it('欄位文字前後有空白仍找得到', function () {
      const r = A.detectColumns([[' 日期 ', ' 星期 ', '項目']]);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
    });

    it('只有日期欄沒有星期欄時視為偵測失敗', function () {
      assertEqual(A.detectColumns([['日期', '溫濕度', 'UPS']]), null);
    });

    it('完全沒有表頭時視為偵測失敗', function () {
      assertEqual(A.detectColumns([['1', '2', '3'], ['4', '5', '6']]), null);
    });

    it('表頭在掃描範圍外時視為偵測失敗', function () {
      const t = [];
      for (let i = 0; i < 12; i++) t.push(['', '', '']);
      t.push(['日期', '星期', '項目']);
      assertEqual(A.detectColumns(t, 10), null);
    });

    it('加大掃描列數後就找得到', function () {
      const t = [];
      for (let i = 0; i < 12; i++) t.push(['', '', '']);
      t.push(['日期', '星期', '項目']);
      assertEqual(A.detectColumns(t, 20).headerRow, 12);
    });
  });
})();
```

- [ ] **Step 2: 把測試檔接上兩個進入點**

`run-tests.js` 加入 `'./tests/table-analysis.test.js'`；`tests.html` 加入：

```html
<script src="src/table-analysis.js"></script>
<script src="tests/table-analysis.test.js"></script>
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/table-analysis.js`

- [ ] **Step 4: 寫實作 `src/table-analysis.js`**

```javascript
// 表格分析：欄位偵測與記號判定，Excel 與 Word 共用
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TableAnalysis = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  // 這些內容視為「沒有資訊」，清空時不需要警示使用者
  const SYMBOLS = ['V', 'v', 'X', 'x', '○', '●', '◯', '✓', '✔', '✗', '✘', '-', '－', 'ˇ'];

  function isSymbolOnly(text) {
    if (text === null || text === undefined) return true;
    const t = String(text).trim();
    if (t === '') return true;
    return SYMBOLS.indexOf(t) !== -1;
  }

  function normalizeHeader(text) {
    return String(text === null || text === undefined ? '' : text).trim();
  }

  function findDateColumn(row) {
    // 先找完整的「日期」，找不到才接受單獨的「日」
    for (let i = 0; i < row.length; i++) {
      if (normalizeHeader(row[i]).indexOf('日期') !== -1) return i;
    }
    for (let i = 0; i < row.length; i++) {
      if (normalizeHeader(row[i]) === '日') return i;
    }
    return -1;
  }

  function findWeekColumn(row) {
    const keywords = ['星期', '週', '禮拜'];
    for (let i = 0; i < row.length; i++) {
      const cell = normalizeHeader(row[i]);
      for (let k = 0; k < keywords.length; k++) {
        if (cell.indexOf(keywords[k]) !== -1) return i;
      }
    }
    return -1;
  }

  function detectColumns(table, maxScanRows) {
    const limit = Math.min(table.length, maxScanRows || 10);
    for (let r = 0; r < limit; r++) {
      const row = table[r] || [];
      const dateCol = findDateColumn(row);
      const weekCol = findWeekColumn(row);
      if (dateCol === -1 || weekCol === -1 || dateCol === weekCol) continue;

      const checkCols = [];
      for (let c = 0; c < row.length; c++) {
        if (c !== dateCol && c !== weekCol) checkCols.push(c);
      }
      return { headerRow: r, dateCol: dateCol, weekCol: weekCol, checkCols: checkCols };
    }
    return null;
  }

  return {
    SYMBOLS: SYMBOLS,
    isSymbolOnly: isSymbolOnly,
    detectColumns: detectColumns
  };
});
```

- [ ] **Step 5: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 17 項全數通過

- [ ] **Step 6: 提交**

```bash
git add src/table-analysis.js tests/table-analysis.test.js run-tests.js tests.html
git commit -m "新增表格分析：日期與星期欄位偵測、勾記符號判定

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 7: 寫入計畫產生器

本專案的業務規則核心。純函式，不碰檔案格式，決定「每一格該寫什麼」，Excel 與 Word 共用同一份判斷。

**Files:**
- Create: `src/write-plan.js`
- Create: `tests/write-plan.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: `TableAnalysis.isSymbolOnly`、`MonthModel.DayModel`
- Produces:
  - `CellWrite = {row: number, col: number, value: string}` — `row`、`col` 皆為 `table` 的索引
  - `buildWritePlan(input): WritePlanResult`
    - `input = {monthModel: DayModel[], table: string[][], headerRow: number, dateCol: number, weekCol: number, checkCols: number[]}`
    - `WritePlanResult = {writes: CellWrite[], rowSummaries: RowSummary[], missingDays: number[], clearedRows: number[], contentLoss: Array<{row, col, text}>}`
    - `RowSummary = {row, day, weekday, isHoliday, isMakeup, description, kind}`，`kind` 為 `'weekday'`｜`'holiday'`｜`'makeup'`｜`'cleared'`
  - 瀏覽器全域變數名：`WritePlan`

**規則**（對應 spec 第三節寫入規則表）：
- 資料列從 `headerRow + 1` 開始，依序對應當月第 1 天、第 2 天……
- 平日與補班日：日期欄寫入天數、星期欄寫入星期、檢查欄一律寫入空字串
- 假日：日期欄與星期欄照寫，檢查欄一律寫入 `-`
- 超出當月天數的列：三種欄位全部寫入空字串，並記入 `clearedRows`
- 表格列數不足以容納當月天數時，容不下的日子記入 `missingDays`
- 凡即將被寫入空字串或 `-`、而原內容不是記號者，記入 `contentLoss`

- [ ] **Step 1: 寫失敗測試 `tests/write-plan.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const W = isNode ? require('../src/write-plan.js') : self.WritePlan;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  // 三天的月份模型：平日、假日、補班日各一
  const 模型 = [
    { day: 1, weekday: '五', isHoliday: false, isMakeup: false, description: '' },
    { day: 2, weekday: '六', isHoliday: true, isMakeup: false, description: '' },
    { day: 3, weekday: '日', isHoliday: false, isMakeup: true, description: '補行上班' }
  ];

  const 欄位 = { headerRow: 0, dateCol: 0, weekCol: 1, checkCols: [2, 3] };

  function 建表(資料列) {
    return [['日期', '星期', '溫濕度', '備註']].concat(資料列);
  }

  function 取值(result, row, col) {
    const hit = result.writes.filter(function (w) { return w.row === row && w.col === col; });
    return hit.length ? hit[0].value : undefined;
  }

  describe('寫入計畫 › 日期與星期', function () {
    const r = W.buildWritePlan(Object.assign({
      monthModel: 模型,
      table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, 欄位));

    it('日期欄依序寫入天數', function () {
      assertEqual(取值(r, 1, 0), '1');
      assertEqual(取值(r, 2, 0), '2');
      assertEqual(取值(r, 3, 0), '3');
    });

    it('星期欄寫入對應星期', function () {
      assertEqual(取值(r, 1, 1), '五');
      assertEqual(取值(r, 2, 1), '六');
      assertEqual(取值(r, 3, 1), '日');
    });
  });

  describe('寫入計畫 › 檢查欄', function () {
    const r = W.buildWritePlan(Object.assign({
      monthModel: 模型,
      table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, 欄位));

    it('平日的檢查欄清空', function () {
      assertEqual(取值(r, 1, 2), '');
      assertEqual(取值(r, 1, 3), '');
    });

    it('假日的檢查欄填入減號', function () {
      assertEqual(取值(r, 2, 2), '-');
      assertEqual(取值(r, 2, 3), '-');
    });

    it('補班日視同平日，檢查欄清空而非填減號', function () {
      assertEqual(取值(r, 3, 2), '');
      assertEqual(取值(r, 3, 3), '');
    });
  });

  describe('寫入計畫 › 列摘要', function () {
    const r = W.buildWritePlan(Object.assign({
      monthModel: 模型,
      table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
    }, 欄位));

    it('每個資料列都有一筆摘要', function () {
      assertEqual(r.rowSummaries.length, 3);
    });

    it('摘要標示出列的種類', function () {
      assertEqual(r.rowSummaries[0].kind, 'weekday');
      assertEqual(r.rowSummaries[1].kind, 'holiday');
      assertEqual(r.rowSummaries[2].kind, 'makeup');
    });

    it('摘要帶出假日名稱', function () {
      assertEqual(r.rowSummaries[2].description, '補行上班');
    });
  });

  describe('寫入計畫 › 多餘與不足的列', function () {
    it('超出天數的列整列清空並記錄', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['9', '一', 'V', 'V']])
      }, 欄位));
      assertDeepEqual(r.clearedRows, [4]);
      assertEqual(取值(r, 4, 0), '');
      assertEqual(取值(r, 4, 1), '');
      assertEqual(取值(r, 4, 2), '');
      assertEqual(r.rowSummaries[3].kind, 'cleared');
    });

    it('列數不足時記錄寫不下的日子', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', '', '']])
      }, 欄位));
      assertDeepEqual(r.missingDays, [2, 3]);
    });

    it('列數剛好時沒有多餘也沒有不足', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', '']])
      }, 欄位));
      assertDeepEqual(r.missingDays, []);
      assertDeepEqual(r.clearedRows, []);
    });
  });

  describe('寫入計畫 › 文字內容警示', function () {
    it('備註欄有真正文字時列入警示', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', 'V', '空調異音已通報'], ['', '', '', ''], ['', '', '', '']])
      }, 欄位));
      assertEqual(r.contentLoss.length, 1);
      assertEqual(r.contentLoss[0].row, 1);
      assertEqual(r.contentLoss[0].col, 3);
      assertEqual(r.contentLoss[0].text, '空調異音已通報');
    });

    it('勾記符號不列入警示', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', 'V', 'X'], ['', '', '○', '-'], ['', '', '', '']])
      }, 欄位));
      assertDeepEqual(r.contentLoss, []);
    });

    it('被整列清空的多餘列也會檢查文字', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['9', '一', 'V', '前月殘留紀錄']])
      }, 欄位));
      assertEqual(r.contentLoss.length, 1);
      assertEqual(r.contentLoss[0].text, '前月殘留紀錄');
    });

    it('日期欄與星期欄的原內容不列入警示', function () {
      const r = W.buildWritePlan(Object.assign({
        monthModel: 模型,
        table: 建表([['31', '三', '', ''], ['', '', '', ''], ['', '', '', '']])
      }, 欄位));
      assertDeepEqual(r.contentLoss, []);
    });
  });
})();
```

- [ ] **Step 2: 把測試檔接上兩個進入點**

`run-tests.js` 加入 `'./tests/write-plan.test.js'`；`tests.html` 加入：

```html
<script src="src/write-plan.js"></script>
<script src="tests/write-plan.test.js"></script>
```

- [ ] **Step 3: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/write-plan.js`

- [ ] **Step 4: 寫實作 `src/write-plan.js`**

```javascript
// 寫入計畫：決定每一格該寫什麼，是本工具的業務規則核心
// 不涉及任何檔案格式，Excel 與 Word 共用同一份判斷
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./table-analysis.js'));
  } else {
    root.WritePlan = factory(root.TableAnalysis);
  }
})(typeof self !== 'undefined' ? self : this, function (TableAnalysis) {
  const HOLIDAY_MARK = '-';

  function cellText(table, row, col) {
    const r = table[row];
    if (!r) return '';
    const v = r[col];
    return v === null || v === undefined ? '' : String(v);
  }

  function buildWritePlan(input) {
    const monthModel = input.monthModel;
    const table = input.table;
    const headerRow = input.headerRow;
    const dateCol = input.dateCol;
    const weekCol = input.weekCol;
    const checkCols = input.checkCols || [];

    const writes = [];
    const rowSummaries = [];
    const clearedRows = [];
    const contentLoss = [];

    const firstDataRow = headerRow + 1;
    const dataRowCount = Math.max(0, table.length - firstDataRow);

    // 記錄一格的寫入，並在覆蓋掉真正的文字時提出警示
    function write(row, col, value, checkLoss) {
      if (checkLoss) {
        const original = cellText(table, row, col);
        if (!TableAnalysis.isSymbolOnly(original)) {
          contentLoss.push({ row: row, col: col, text: original });
        }
      }
      writes.push({ row: row, col: col, value: value });
    }

    for (let i = 0; i < dataRowCount; i++) {
      const row = firstDataRow + i;
      const day = monthModel[i];

      if (!day) {
        // 超出當月天數，整列清空
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

      const mark = day.isHoliday ? HOLIDAY_MARK : '';
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

    // 表格列數不足時，記錄容納不下的日子
    const missingDays = [];
    for (let i = dataRowCount; i < monthModel.length; i++) {
      missingDays.push(monthModel[i].day);
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
```

- [ ] **Step 5: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 15 項全數通過

- [ ] **Step 6: 提交**

```bash
git add src/write-plan.js tests/write-plan.test.js run-tests.js tests.html
git commit -m "新增寫入計畫產生器，集中所有寫入業務規則

平日清空、假日填減號、補班日視同平日、多餘列整列清空，
並在覆蓋真正的文字時提出警示。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 8: 仿真範例檔

使用者手邊沒有可用的檢查表，必須自行產生測試素材。以真實的 Office 函式庫產生，而非手寫 XML——手寫出來的只是「我們以為的格式」，拿它測自己的解析器等於自我驗證，沒有意義。

**Files:**
- Create: `package.json`
- Create: `tools/make-samples.js`
- Create: `samples/機房月檢查表-範例.xlsx`（由腳本產生）
- Create: `samples/機房月檢查表-範例.docx`（由腳本產生）

**Interfaces:**
- Consumes: 無
- Produces: 兩個範例檔，供 Task 9 至 12 與 Task 15 使用

**這些 npm 套件只在開發端使用，絕不進入發布檔。**

- [ ] **Step 1: 建立 `package.json`**

```json
{
  "name": "week-automodify",
  "version": "1.0.0",
  "description": "機房月檢查表星期自動填寫工具",
  "private": true,
  "scripts": {
    "test": "node run-tests.js",
    "samples": "node tools/make-samples.js",
    "holidays": "node tools/fetch-holidays.js",
    "build": "node build.js"
  },
  "devDependencies": {
    "docx": "^9.0.0",
    "exceljs": "^4.4.0",
    "jszip": "^3.10.1"
  }
}
```

- [ ] **Step 2: 安裝開發相依套件**

Run: `npm install`
Expected: 安裝完成，產生 `node_modules/`（已由 `.gitignore` 排除）

- [ ] **Step 3: 把 JSZip 複製為發布用的本地副本**

發布檔不能連 CDN，因此需要一份本地副本。

Run:
```bash
mkdir -p vendor
node -e "const fs=require('fs');fs.copyFileSync('node_modules/jszip/dist/jszip.min.js','vendor/jszip.min.js');console.log('已複製 JSZip');"
```
Expected: 產生 `vendor/jszip.min.js`，約 96 KB

- [ ] **Step 4: 寫範例檔產生腳本 `tools/make-samples.js`**

範例刻意包含幾種難處理的情形：標題含年月、表頭前有空白列、備註欄有真正的文字、表格固定 31 列（比小月多）。

```javascript
// 開發端工具：產生仿真的機房月檢查表，供測試使用
// 用法：node tools/make-samples.js
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } = require('docx');

const 表頭 = ['日期', '星期', '溫濕度', 'UPS', '消防設備', '門禁', '備註'];
const 標題 = '機房狀態月檢查表';
const 年月 = '115年8月';
const 總列數 = 31;
const outDir = path.join(__dirname, '..', 'samples');

// 第 3 列的備註欄故意放真正的文字，用來驗證文字內容警示
function 資料列(day) {
  const 備註 = day === 3 ? '空調異音已通報維護廠商' : '';
  return [String(day), '', 'V', 'V', 'V', 'V', 備註];
}

async function 產生Excel() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('8月');

  ws.mergeCells(1, 1, 1, 表頭.length);
  ws.getCell(1, 1).value = 標題 + '  ' + 年月;
  ws.getCell(1, 1).alignment = { horizontal: 'center' };
  ws.getCell(1, 1).font = { size: 14, bold: true };

  ws.addRow([]); // 第 2 列刻意留白，用來驗證表頭偵測不是寫死在第一列
  ws.addRow(表頭);
  for (let d = 1; d <= 總列數; d++) ws.addRow(資料列(d));

  // 表頭與資料列加框線，用來驗證改寫後框線不會消失
  const 細框 = { style: 'thin' };
  for (let r = 3; r <= 3 + 總列數; r++) {
    for (let c = 1; c <= 表頭.length; c++) {
      ws.getCell(r, c).border = { top: 細框, left: 細框, bottom: 細框, right: 細框 };
    }
  }
  ws.getRow(3).font = { bold: true };

  const p = path.join(outDir, '機房月檢查表-範例.xlsx');
  await wb.xlsx.writeFile(p);
  console.log('已產生 ' + p);
}

function 表格列(cells, 粗體) {
  return new TableRow({
    children: cells.map(function (t) {
      return new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: t, bold: !!粗體 })] })]
      });
    })
  });
}

async function 產生Word() {
  const rows = [表格列(表頭, true)];
  for (let d = 1; d <= 總列數; d++) rows.push(表格列(資料列(d), false));

  const doc = new Document({
    sections: [{
      children: [
        // 標題刻意拆成兩個 run，模擬 Word 把一段文字切碎的實際情形
        new Paragraph({
          children: [
            new TextRun(標題 + '  '),
            new TextRun({ text: 年月, bold: true })
          ]
        }),
        new Paragraph(''),
        new Table({ rows: rows })
      ]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  const p = path.join(outDir, '機房月檢查表-範例.docx');
  fs.writeFileSync(p, buf);
  console.log('已產生 ' + p);
}

(async function () {
  fs.mkdirSync(outDir, { recursive: true });
  await 產生Excel();
  await 產生Word();
})();
```

- [ ] **Step 5: 執行腳本產生範例檔**

Run: `node tools/make-samples.js`
Expected: 印出兩個檔案路徑，`samples/` 下出現兩個檔案

- [ ] **Step 6: 確認範例檔的內部結構符合預期**

這一步是要親眼確認 Task 9 至 12 將面對的實際結構。

Run:
```bash
node -e "
const JSZip=require('jszip'),fs=require('fs');
(async()=>{
  const z=await JSZip.loadAsync(fs.readFileSync('samples/機房月檢查表-範例.xlsx'));
  const s=await z.file('xl/sharedStrings.xml').async('string');
  console.log('sharedStrings 含年月：', s.indexOf('115年8月')!==-1);
  const sheet=await z.file('xl/worksheets/sheet1.xml').async('string');
  console.log('工作表含 row r=\"3\"：', sheet.indexOf('<row r=\"3\"')!==-1);
  const w=await JSZip.loadAsync(fs.readFileSync('samples/機房月檢查表-範例.docx'));
  const d=await w.file('word/document.xml').async('string');
  console.log('Word 標題被切成多個 run：', d.indexOf('115年8月')!==-1 && d.indexOf('</w:r><w:r>')!==-1);
  console.log('Word 表格列數：', (d.match(/<w:tr>/g)||[]).length);
})();
"
```
Expected: 三項皆為 `true`，Word 表格列數為 32（表頭一列加 31 個資料列）

- [ ] **Step 7: 提交**

範例檔雖是產物，但必須進版控——它們是端對端測試的依據，且使用者端無法自行產生。

```bash
git add package.json tools/make-samples.js vendor/jszip.min.js samples/
git commit -m "新增仿真範例檔與其產生腳本

以真實 Office 函式庫產生，範例刻意包含標題年月、表頭前空白列、
備註欄文字、以及多於小月天數的 31 列，供端對端測試使用。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---
### Task 9: Excel 讀取與解析

把 `.xlsx` 解開成二維字串表格。**不得使用 DOM 解析器**——Node 沒有內建 `DOMParser`，而本模組必須在兩種環境都能跑，因此一律以正規式處理 XML。

**Files:**
- Create: `src/xml-text.js`
- Create: `src/excel-adapter.js`
- Create: `tests/xml-text.test.js`
- Create: `tests/excel-adapter.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: JSZip（以參數注入，Node 端用 `require('jszip')`，瀏覽器端用全域 `JSZip`）
- Produces:
  - `XmlText.unescape(s: string): string`、`XmlText.escape(s: string): string`
  - `ExcelAdapter.colToIndex(letters: string): number` — `'A'` → 0、`'AA'` → 26
  - `ExcelAdapter.indexToCol(index: number): string` — 0 → `'A'`
  - `ExcelAdapter.parseCellRef(ref: string): {row: number, col: number}` — `'B5'` → `{row: 4, col: 1}`（皆為 0 起算）
  - `ExcelAdapter.parseSharedStrings(xml: string): string[]`
  - `ExcelAdapter.parseSheet(sheetXml: string, sharedStrings: string[]): {table: string[][], cells: CellInfo[]}`
    - `CellInfo = {ref, row, col, styleAttr, type, raw}`，`styleAttr` 為原本的 `s="n"` 字串（沒有則為空字串）
  - `ExcelAdapter.loadWorkbook(data, JSZipLib): Promise<{zip, sheetPath, sheetXml, sharedStringsXml, sharedStrings, table, cells}>`
  - 瀏覽器全域變數名：`XmlText`、`ExcelAdapter`

**實測得知的關鍵事實**（來自對真實 `.xlsx` 的檢視）：
- 空白列在 XML 中**完全不存在**（`<row r="1">` 之後直接是 `<row r="3">`），因此**必須依 `r` 屬性定位，不可依出現順序推算**
- 文字一律存於 `xl/sharedStrings.xml`，儲存格內只放索引：`<c r="B4" s="1" t="s"><v>7</v></c>`
- 樣式存於 `s` 屬性，與 `t` 屬性完全分離，因此改變值不會動到框線
- 空儲存格寫作自閉合標籤：`<c r="B1"/>`

- [ ] **Step 1: 寫 XML 跳脫的失敗測試 `tests/xml-text.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const X = isNode ? require('../src/xml-text.js') : self.XmlText;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  describe('XML 文字處理', function () {
    it('還原五種實體', function () {
      assertEqual(X.unescape('a&amp;b&lt;c&gt;d&quot;e&apos;f'), 'a&b<c>d"e\'f');
    });

    it('跳脫五種字元', function () {
      assertEqual(X.escape('a&b<c>d"e\'f'), 'a&amp;b&lt;c&gt;d&quot;e&apos;f');
    });

    it('還原數值實體', function () {
      assertEqual(X.unescape('&#65;&#x42;'), 'AB');
    });

    it('中文不受影響', function () {
      assertEqual(X.escape('機房檢查'), '機房檢查');
      assertEqual(X.unescape('機房檢查'), '機房檢查');
    });

    it('跳脫後再還原可回到原文', function () {
      const s = '備註 <重要> & "緊急"';
      assertEqual(X.unescape(X.escape(s)), s);
    });

    it('先還原 amp 會出錯，順序必須正確', function () {
      // &amp;lt; 應還原為 &lt; 這個字面文字，而非 <
      assertEqual(X.unescape('&amp;lt;'), '&lt;');
    });

    it('空字串與 null 安全', function () {
      assertEqual(X.escape(''), '');
      assertEqual(X.unescape(''), '');
      assertEqual(X.escape(null), '');
      assertEqual(X.unescape(null), '');
    });
  });
})();
```

- [ ] **Step 2: 寫實作 `src/xml-text.js`**

還原的順序很重要：`&amp;` 必須**最後**才還原，否則 `&amp;lt;` 會被錯誤地變成 `<`。

```javascript
// XML 文字的跳脫與還原
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XmlText = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  function escape(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function unescape(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, hex) {
        return String.fromCodePoint(parseInt(hex, 16));
      })
      .replace(/&#([0-9]+);/g, function (_, dec) {
        return String.fromCodePoint(parseInt(dec, 10));
      })
      // &amp; 必須最後還原，否則 &amp;lt; 會被誤判成 <
      .replace(/&amp;/g, '&');
  }

  return { escape: escape, unescape: unescape };
});
```

- [ ] **Step 3: 接上進入點並確認測試通過**

`run-tests.js` 加入 `'./tests/xml-text.test.js'`；`tests.html` 加入 `src/xml-text.js` 與該測試檔。

Run: `node run-tests.js`
Expected: PASS，新增的 7 項全數通過

- [ ] **Step 4: 寫 Excel 解析的失敗測試 `tests/excel-adapter.test.js`**

前半部用寫死的 XML 片段測純函式，後半部讀入真實範例檔做整合測試。

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const E = isNode ? require('../src/excel-adapter.js') : self.ExcelAdapter;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('Excel › 欄名轉換', function () {
    it('單字母欄名', function () {
      assertEqual(E.colToIndex('A'), 0);
      assertEqual(E.colToIndex('B'), 1);
      assertEqual(E.colToIndex('Z'), 25);
    });

    it('雙字母欄名', function () {
      assertEqual(E.colToIndex('AA'), 26);
      assertEqual(E.colToIndex('AB'), 27);
      assertEqual(E.colToIndex('BA'), 52);
    });

    it('索引轉回欄名', function () {
      assertEqual(E.indexToCol(0), 'A');
      assertEqual(E.indexToCol(25), 'Z');
      assertEqual(E.indexToCol(26), 'AA');
      assertEqual(E.indexToCol(27), 'AB');
    });

    it('來回轉換一致', function () {
      for (let i = 0; i < 100; i++) assertEqual(E.colToIndex(E.indexToCol(i)), i);
    });

    it('儲存格參照拆解為零起算的列與欄', function () {
      assertDeepEqual(E.parseCellRef('A1'), { row: 0, col: 0 });
      assertDeepEqual(E.parseCellRef('B5'), { row: 4, col: 1 });
      assertDeepEqual(E.parseCellRef('AA10'), { row: 9, col: 26 });
    });
  });

  describe('Excel › 共用字串解析', function () {
    it('取出所有字串', function () {
      const xml = '<sst count="3" uniqueCount="3"><si><t>日期</t></si>' +
        '<si><t>星期</t></si><si><t>備註</t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['日期', '星期', '備註']);
    });

    it('空字串項目保留為空字串', function () {
      const xml = '<sst><si><t>甲</t></si><si><t></t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['甲', '']);
    });

    it('還原跳脫字元', function () {
      const xml = '<sst><si><t>A&amp;B</t></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['A&B']);
    });

    it('處理被切成多段的富文字', function () {
      const xml = '<sst><si><r><t>115年</t></r><r><t>8月</t></r></si></sst>';
      assertDeepEqual(E.parseSharedStrings(xml), ['115年8月']);
    });

    it('沒有共用字串時回傳空陣列', function () {
      assertDeepEqual(E.parseSharedStrings('<sst/>'), []);
    });
  });

  describe('Excel › 工作表解析', function () {
    const 共用 = ['標題', '日期', '星期', '備註', '六'];
    const 工作表 =
      '<worksheet><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' +
      '<row r="3"><c r="A3" s="1" t="s"><v>1</v></c>' +
      '<c r="B3" s="1" t="s"><v>2</v></c><c r="C3" s="1" t="s"><v>3</v></c></row>' +
      '<row r="4"><c r="A4" s="1"><v>1</v></c>' +
      '<c r="B4" s="1" t="s"><v>4</v></c><c r="C4" s="1"/></row>' +
      '</sheetData></worksheet>';

    it('依 r 屬性定位，跳過的列成為空列', function () {
      const r = E.parseSheet(工作表, 共用);
      assertEqual(r.table.length, 4, '應有第 1 至第 4 列');
      assertDeepEqual(r.table[1], [], '第 2 列在 XML 中不存在，應為空列');
    });

    it('共用字串型的儲存格取出文字而非索引', function () {
      const r = E.parseSheet(工作表, 共用);
      assertEqual(r.table[2][0], '日期');
      assertEqual(r.table[2][1], '星期');
    });

    it('數字型儲存格取出數字文字', function () {
      assertEqual(E.parseSheet(工作表, 共用).table[3][0], '1');
    });

    it('自閉合的空儲存格為空字串', function () {
      assertEqual(E.parseSheet(工作表, 共用).table[3][2], '');
    });

    it('保留每一格原本的樣式屬性', function () {
      const r = E.parseSheet(工作表, 共用);
      const b4 = r.cells.filter(function (c) { return c.ref === 'B4'; })[0];
      assertEqual(b4.styleAttr, ' s="1"');
      assertEqual(b4.row, 3);
      assertEqual(b4.col, 1);
    });

    it('行內字串型的儲存格也讀得到', function () {
      const xml = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>行內</t></is></c></row></sheetData>';
      assertEqual(E.parseSheet(xml, []).table[0][0], '行內');
    });
  });

  describe('Excel › 讀取真實範例檔', function () {
    // 此區塊需要檔案系統，僅在 Node 端執行
    if (!isNode) return;
    const fs = require('fs');
    const JSZip = require('jszip');
    const A = require('../src/table-analysis.js');

    it('讀得出表格內容', async function () {
      const buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      const wb = await E.loadWorkbook(buf, JSZip);
      assertEqual(wb.table[0][0], '機房狀態月檢查表  115年8月', '第 1 列為標題');
      assertEqual(wb.table[2][0], '日期', '第 3 列為表頭');
      assertEqual(wb.table[2][1], '星期');
      assertEqual(wb.table[3][0], '1', '第 4 列為第一天');
    });

    it('欄位偵測能在真實檔案上成功', async function () {
      const buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      const wb = await E.loadWorkbook(buf, JSZip);
      const cols = A.detectColumns(wb.table);
      assertEqual(cols.headerRow, 2);
      assertEqual(cols.dateCol, 0);
      assertEqual(cols.weekCol, 1);
      assertEqual(cols.checkCols.length, 5);
    });

    it('備註欄的文字讀得到', async function () {
      const buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      const wb = await E.loadWorkbook(buf, JSZip);
      assertEqual(wb.table[5][6], '空調異音已通報維護廠商', '第 3 天的備註');
    });
  });
})();
```

- [ ] **Step 5: 接上進入點並確認測試失敗**

`run-tests.js` 加入 `'./tests/excel-adapter.test.js'`；`tests.html` 加入 `src/excel-adapter.js` 與該測試檔。

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/excel-adapter.js`

- [ ] **Step 6: 寫實作 `src/excel-adapter.js`**

```javascript
// Excel 轉接層：以正規式直接處理 xlsx 內部 XML，不使用 DOM 解析器
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'));
  } else {
    root.ExcelAdapter = factory(root.XmlText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText) {

  function colToIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n - 1;
  }

  function indexToCol(index) {
    let n = index + 1;
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function parseCellRef(ref) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) return null;
    return { row: parseInt(m[2], 10) - 1, col: colToIndex(m[1]) };
  }

  function parseSharedStrings(xml) {
    const out = [];
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let m;
    while ((m = siRe.exec(xml)) !== null) {
      const inner = m[1] || '';
      // 一個 si 可能被切成多個 t，需全部串接
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
      let t;
      let text = '';
      while ((t = tRe.exec(inner)) !== null) {
        text += XmlText.unescape(t[1] || '');
      }
      out.push(text);
    }
    return out;
  }

  function cellValue(attrs, inner, sharedStrings) {
    const typeMatch = /\bt="([^"]*)"/.exec(attrs);
    const type = typeMatch ? typeMatch[1] : '';

    if (type === 'inlineStr') {
      let text = '';
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let t;
      while ((t = tRe.exec(inner)) !== null) text += XmlText.unescape(t[1]);
      return { type: type, text: text };
    }

    const vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
    const rawValue = vMatch ? XmlText.unescape(vMatch[1]) : '';

    if (type === 's') {
      const idx = parseInt(rawValue, 10);
      return { type: type, text: sharedStrings[idx] === undefined ? '' : sharedStrings[idx] };
    }
    return { type: type, text: rawValue };
  }

  function parseSheet(sheetXml, sharedStrings) {
    const table = [];
    const cells = [];
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
    let rowMatch;

    while ((rowMatch = rowRe.exec(sheetXml)) !== null) {
      const rowAttrs = rowMatch[1] !== undefined ? rowMatch[1] : rowMatch[3];
      const rowInner = rowMatch[2] || '';
      const rMatch = /\br="([0-9]+)"/.exec(rowAttrs);
      if (!rMatch) continue;
      // 依 r 屬性定位，因為空白列在 XML 中根本不存在
      const rowIndex = parseInt(rMatch[1], 10) - 1;

      while (table.length <= rowIndex) table.push([]);
      const rowArray = table[rowIndex];

      const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g;
      let cellMatch;
      while ((cellMatch = cellRe.exec(rowInner)) !== null) {
        const attrs = cellMatch[1] !== undefined ? cellMatch[1] : cellMatch[3];
        const inner = cellMatch[2] || '';
        const refMatch = /\br="([A-Z]+[0-9]+)"/.exec(attrs);
        if (!refMatch) continue;
        const ref = refMatch[1];
        const pos = parseCellRef(ref);
        const styleMatch = /\s(s="[0-9]+")/.exec(attrs);
        const value = cellValue(attrs, inner, sharedStrings);

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

  // 找出活頁簿中第一個工作表的路徑
  function firstSheetPath(zip) {
    const names = Object.keys(zip.files).filter(function (n) {
      return /^xl\/worksheets\/sheet[0-9]+\.xml$/.test(n);
    });
    names.sort();
    return names[0] || null;
  }

  function loadWorkbook(data, JSZipLib) {
    return JSZipLib.loadAsync(data).then(function (zip) {
      const sheetPath = firstSheetPath(zip);
      if (!sheetPath) throw new Error('這個檔案裡找不到工作表，可能不是有效的 Excel 檔');

      const ssFile = zip.file('xl/sharedStrings.xml');
      const jobs = [zip.file(sheetPath).async('string')];
      jobs.push(ssFile ? ssFile.async('string') : Promise.resolve(''));

      return Promise.all(jobs).then(function (parts) {
        const sheetXml = parts[0];
        const sharedStringsXml = parts[1];
        const sharedStrings = sharedStringsXml ? parseSharedStrings(sharedStringsXml) : [];
        const parsed = parseSheet(sheetXml, sharedStrings);
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

  return {
    colToIndex: colToIndex,
    indexToCol: indexToCol,
    parseCellRef: parseCellRef,
    parseSharedStrings: parseSharedStrings,
    parseSheet: parseSheet,
    loadWorkbook: loadWorkbook
  };
});
```

- [ ] **Step 7: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 19 項全數通過

- [ ] **Step 8: 提交**

```bash
git add src/xml-text.js src/excel-adapter.js tests/xml-text.test.js tests/excel-adapter.test.js run-tests.js tests.html
git commit -m "新增 Excel 讀取與 XML 文字處理

以正規式處理 xlsx 內部 XML 以兼顧 Node 與瀏覽器，
並依 r 屬性定位儲存格，因為空白列在 XML 中並不存在。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 10: Excel 寫入

把寫入計畫套回 xlsx，並改寫年月文字。**格式必須毫髮無傷**。

**Files:**
- Modify: `src/excel-adapter.js`（新增寫入相關函式）
- Modify: `tests/excel-adapter.test.js`（新增寫入測試）

**Interfaces:**
- Consumes: `WritePlan.CellWrite`、`YearMonthText`
- Produces（皆掛在 `ExcelAdapter` 上）：
  - `buildCellXml(ref, styleAttr, value): string` — 值為空字串時產生自閉合標籤
  - `applyWritesToSheet(sheetXml, writes, cells): {xml: string, skipped: CellWrite[]}`
  - `findYearMonthTargets(sharedStrings): Array<{index, text, matches}>`
  - `rewriteSharedStrings(xml, sharedStrings, targets, targetYear, targetMonth, selectedKeys): string`
  - `saveWorkbook(zip, sheetPath, sheetXml, sharedStringsXml, JSZipLib, outputType): Promise<Blob|Buffer>`

**年月改寫的做法**：文字都存在 `sharedStrings.xml`，因此直接改該檔中 `<t>` 的內容即可。**字串的數量沒有改變，`count` 與 `uniqueCount` 不需更動**。同一字串若被多處引用，改一次即全部生效，這正是所要的結果。

**清空儲存格的做法**：保留 `<c>` 與其 `s` 樣式屬性，移除 `t` 屬性與內部元素，寫成 `<c r="C4" s="1"/>`。這是 Excel 自己也在用的寫法，框線與底色不受影響。

- [ ] **Step 1: 在 `tests/excel-adapter.test.js` 末尾追加寫入測試**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const E = isNode ? require('../src/excel-adapter.js') : self.ExcelAdapter;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  describe('Excel › 儲存格 XML 產生', function () {
    it('有值時產生行內字串', function () {
      assertEqual(E.buildCellXml('B4', ' s="1"', '六'),
        '<c r="B4" s="1" t="inlineStr"><is><t>六</t></is></c>');
    });

    it('空值時產生自閉合標籤並保留樣式', function () {
      assertEqual(E.buildCellXml('C4', ' s="1"', ''), '<c r="C4" s="1"/>');
    });

    it('沒有樣式屬性時也能產生', function () {
      assertEqual(E.buildCellXml('A1', '', 'X'),
        '<c r="A1" t="inlineStr"><is><t>X</t></is></c>');
    });

    it('值中的特殊字元會被跳脫', function () {
      assertEqual(E.buildCellXml('A1', '', 'a<b&c'),
        '<c r="A1" t="inlineStr"><is><t>a&lt;b&amp;c</t></is></c>');
    });

    it('減號可正常寫入', function () {
      assertEqual(E.buildCellXml('C5', ' s="2"', '-'),
        '<c r="C5" s="2" t="inlineStr"><is><t>-</t></is></c>');
    });
  });

  describe('Excel › 套用寫入', function () {
    const 共用 = ['日期', '星期', '六'];
    const 工作表 =
      '<worksheet><sheetData>' +
      '<row r="3"><c r="A3" s="1" t="s"><v>0</v></c><c r="B3" s="1" t="s"><v>1</v></c></row>' +
      '<row r="4"><c r="A4" s="1"><v>1</v></c><c r="B4" s="1" t="s"><v>2</v></c>' +
      '<c r="C4" s="1"><v>99</v></c></row>' +
      '</sheetData></worksheet>';

    it('改寫後的值讀得回來', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [{ row: 3, col: 1, value: '日' }], parsed.cells);
      const 再讀 = E.parseSheet(r.xml, 共用);
      assertEqual(再讀.table[3][1], '日');
    });

    it('清空後該格為空字串', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [{ row: 3, col: 2, value: '' }], parsed.cells);
      assertEqual(E.parseSheet(r.xml, 共用).table[3][2], '');
    });

    it('樣式屬性完整保留', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [{ row: 3, col: 2, value: '' }], parsed.cells);
      assertEqual(r.xml.indexOf('<c r="C4" s="1"/>') !== -1, true, '應保留 s="1"');
    });

    it('未被寫入的儲存格原封不動', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [{ row: 3, col: 1, value: '日' }], parsed.cells);
      assertEqual(r.xml.indexOf('<c r="A3" s="1" t="s"><v>0</v></c>') !== -1, true);
    });

    it('目標儲存格不存在時列入略過清單', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [{ row: 3, col: 9, value: 'X' }], parsed.cells);
      assertEqual(r.skipped.length, 1);
      assertEqual(r.skipped[0].col, 9);
    });

    it('一次寫入多格皆生效', function () {
      const parsed = E.parseSheet(工作表, 共用);
      const r = E.applyWritesToSheet(工作表, [
        { row: 3, col: 0, value: '5' },
        { row: 3, col: 1, value: '三' },
        { row: 3, col: 2, value: '-' }
      ], parsed.cells);
      const 再讀 = E.parseSheet(r.xml, 共用);
      assertEqual(再讀.table[3][0], '5');
      assertEqual(再讀.table[3][1], '三');
      assertEqual(再讀.table[3][2], '-');
    });
  });

  describe('Excel › 共用字串中的年月改寫', function () {
    const xml = '<?xml version="1.0"?><sst count="3" uniqueCount="3">' +
      '<si><t>機房狀態月檢查表  115年8月</t></si>' +
      '<si><t>日期</t></si><si><t>製表：115年8月</t></si></sst>';
    const 共用 = ['機房狀態月檢查表  115年8月', '日期', '製表：115年8月'];

    it('找得到所有含年月的字串', function () {
      const targets = E.findYearMonthTargets(共用);
      assertEqual(targets.length, 2);
      assertEqual(targets[0].index, 0);
      assertEqual(targets[1].index, 2);
    });

    it('沒有年月的字串不列入', function () {
      const targets = E.findYearMonthTargets(['日期', '星期']);
      assertEqual(targets.length, 0);
    });

    it('全部改寫後兩處都變成新年月', function () {
      const targets = E.findYearMonthTargets(共用);
      const out = E.rewriteSharedStrings(xml, 共用, targets, 2026, 9, null);
      assertEqual(out.indexOf('115年9月') !== -1, true);
      assertEqual(out.indexOf('115年8月') === -1, true, '舊年月應全部被取代');
    });

    it('只改寫選定的其中一處', function () {
      const targets = E.findYearMonthTargets(共用);
      const out = E.rewriteSharedStrings(xml, 共用, targets, 2026, 9, ['0-0']);
      assertEqual(out.indexOf('機房狀態月檢查表  115年9月') !== -1, true);
      assertEqual(out.indexOf('製表：115年8月') !== -1, true, '未選取的應維持原狀');
    });

    it('count 與 uniqueCount 不被更動', function () {
      const targets = E.findYearMonthTargets(共用);
      const out = E.rewriteSharedStrings(xml, 共用, targets, 2026, 9, null);
      assertEqual(out.indexOf('count="3" uniqueCount="3"') !== -1, true);
    });

    it('改寫後仍解析得回正確字串', function () {
      const targets = E.findYearMonthTargets(共用);
      const out = E.rewriteSharedStrings(xml, 共用, targets, 2026, 12, null);
      assertEqual(E.parseSharedStrings(out)[0], '機房狀態月檢查表  115年12月');
    });
  });

  describe('Excel › 完整改寫真實範例檔', function () {
    if (!isNode) return;
    const fs = require('fs');
    const JSZip = require('jszip');
    const A = require('../src/table-analysis.js');
    const M = require('../src/month-model.js');
    const H = require('../src/holidays-tw.js');
    const B = require('../src/holidays-builtin.js');
    const W = require('../src/write-plan.js');

    async function 產生改寫結果() {
      const buf = fs.readFileSync('samples/機房月檢查表-範例.xlsx');
      const wb = await E.loadWorkbook(buf, JSZip);
      const cols = A.detectColumns(wb.table);
      const model = M.buildMonth(2026, 9, H.normalizeYearData(B['2026']));
      const plan = W.buildWritePlan({
        monthModel: model, table: wb.table, headerRow: cols.headerRow,
        dateCol: cols.dateCol, weekCol: cols.weekCol, checkCols: cols.checkCols
      });
      const applied = E.applyWritesToSheet(wb.sheetXml, plan.writes, wb.cells);
      const targets = E.findYearMonthTargets(wb.sharedStrings);
      const newSs = E.rewriteSharedStrings(wb.sharedStringsXml, wb.sharedStrings, targets, 2026, 9, null);
      const out = await E.saveWorkbook(wb.zip, wb.sheetPath, applied.xml, newSs, JSZip, 'nodebuffer');
      return { out: out, plan: plan, applied: applied };
    }

    it('改寫後的檔案仍可被重新讀取', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[2][0], '日期', '表頭未受影響');
    });

    it('日期與星期正確寫入', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[3][0], '1', '9月1日');
      assertEqual(wb2.table[3][1], '二', '9月1日為星期二');
      assertEqual(wb2.table[7][0], '5');
      assertEqual(wb2.table[7][1], '六', '9月5日為星期六');
    });

    it('假日的檢查欄填入減號', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[7][2], '-', '9月5日為週六');
      assertEqual(wb2.table[27][2], '-', '9月25日為中秋節');
    });

    it('平日的檢查欄被清空', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[3][2], '', '9月1日為平日');
    });

    it('超出當月天數的第 31 列被整列清空', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[33][0], '', '9月只有 30 天，第 31 列應清空');
      assertEqual(wb2.table[33][1], '');
    });

    it('標題的年月被改寫', async function () {
      const r = await 產生改寫結果();
      const wb2 = await E.loadWorkbook(r.out, JSZip);
      assertEqual(wb2.table[0][0], '機房狀態月檢查表  115年9月');
    });

    it('備註欄的文字被列入警示', async function () {
      const r = await 產生改寫結果();
      assertEqual(r.plan.contentLoss.length, 1);
      assertEqual(r.plan.contentLoss[0].text, '空調異音已通報維護廠商');
    });

    it('沒有任何儲存格被略過', async function () {
      const r = await 產生改寫結果();
      assertEqual(r.applied.skipped.length, 0);
    });

    it('框線樣式檔完全未被更動', async function () {
      const r = await 產生改寫結果();
      const 原檔 = await JSZip.loadAsync(fs.readFileSync('samples/機房月檢查表-範例.xlsx'));
      const 新檔 = await JSZip.loadAsync(r.out);
      const a = await 原檔.file('xl/styles.xml').async('string');
      const b = await 新檔.file('xl/styles.xml').async('string');
      assertEqual(a === b, true, 'styles.xml 必須逐字相同');
    });
  });
})();
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，`E.buildCellXml is not a function`

- [ ] **Step 3: 在 `src/excel-adapter.js` 中新增寫入相關函式**

在 `return` 敘述之前插入下列函式，並把它們加進回傳的物件。此模組需要用到 `YearMonthText`，因此工廠函式的相依也要一併調整。

先把模組頭尾改成同時注入兩個相依：

```javascript
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'), require('./year-month-text.js'));
  } else {
    root.ExcelAdapter = factory(root.XmlText, root.YearMonthText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText, YearMonthText) {
```

新增的函式：

```javascript
  // 值為空時產生自閉合標籤，保留樣式屬性以免框線消失
  function buildCellXml(ref, styleAttr, value) {
    const style = styleAttr || '';
    if (value === '' || value === null || value === undefined) {
      return '<c r="' + ref + '"' + style + '/>';
    }
    return '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t>' +
      XmlText.escape(value) + '</t></is></c>';
  }

  function applyWritesToSheet(sheetXml, writes, cells) {
    const byPosition = {};
    cells.forEach(function (c) { byPosition[c.row + ',' + c.col] = c; });

    const skipped = [];
    const replacements = [];

    writes.forEach(function (w) {
      const cell = byPosition[w.row + ',' + w.col];
      if (!cell) { skipped.push(w); return; }
      replacements.push({
        from: cell.raw,
        to: buildCellXml(cell.ref, cell.styleAttr, w.value)
      });
    });

    let xml = sheetXml;
    replacements.forEach(function (r) {
      // 逐一取代原始片段，不使用全域取代以免波及同樣內容的其他儲存格
      const at = xml.indexOf(r.from);
      if (at === -1) return;
      xml = xml.slice(0, at) + r.to + xml.slice(at + r.from.length);
    });

    return { xml: xml, skipped: skipped };
  }

  // 找出共用字串中所有含年月的項目
  function findYearMonthTargets(sharedStrings) {
    const out = [];
    sharedStrings.forEach(function (text, index) {
      const matches = YearMonthText.findYearMonth(text);
      if (matches.length > 0) out.push({ index: index, text: text, matches: matches });
    });
    return out;
  }

  // 直接改寫 sharedStrings.xml 中的文字。字串數量不變，
  // 因此 count 與 uniqueCount 不需更動。
  function rewriteSharedStrings(xml, sharedStrings, targets, targetYear, targetMonth, selectedKeys) {
    let out = xml;

    targets.forEach(function (target) {
      const chosen = [];
      target.matches.forEach(function (m, i) {
        const key = target.index + '-' + i;
        if (!selectedKeys || selectedKeys.indexOf(key) !== -1) chosen.push(i);
      });
      if (chosen.length === 0) return;

      const newText = YearMonthText.applyRewrites(
        target.text, target.matches, targetYear, targetMonth, chosen
      );
      if (newText === target.text) return;

      const oldEscaped = XmlText.escape(target.text);
      const newEscaped = XmlText.escape(newText);
      const at = out.indexOf('>' + oldEscaped + '<');
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
```

把這五個函式加進 `return` 物件。

- [ ] **Step 4: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 27 項全數通過。特別確認「框線樣式檔完全未被更動」這一項通過——它是格式保護的主要證據。

- [ ] **Step 5: 產生一份實際輸出供人工檢視**

Run:
```bash
node -e "
const fs=require('fs'),JSZip=require('jszip');
const E=require('./src/excel-adapter.js'),A=require('./src/table-analysis.js');
const M=require('./src/month-model.js'),H=require('./src/holidays-tw.js');
const B=require('./src/holidays-builtin.js'),W=require('./src/write-plan.js');
(async()=>{
  const wb=await E.loadWorkbook(fs.readFileSync('samples/機房月檢查表-範例.xlsx'),JSZip);
  const cols=A.detectColumns(wb.table);
  const model=M.buildMonth(2026,9,H.normalizeYearData(B['2026']));
  const plan=W.buildWritePlan({monthModel:model,table:wb.table,headerRow:cols.headerRow,dateCol:cols.dateCol,weekCol:cols.weekCol,checkCols:cols.checkCols});
  const applied=E.applyWritesToSheet(wb.sheetXml,plan.writes,wb.cells);
  const targets=E.findYearMonthTargets(wb.sharedStrings);
  const ss=E.rewriteSharedStrings(wb.sharedStringsXml,wb.sharedStrings,targets,2026,9,null);
  const out=await E.saveWorkbook(wb.zip,wb.sheetPath,applied.xml,ss,JSZip,'nodebuffer');
  fs.writeFileSync('samples/輸出檢視-115年09月.xlsx',out);
  console.log('已產生 samples/輸出檢視-115年09月.xlsx');
  const wb2=await E.loadWorkbook(out,JSZip);
  wb2.table.slice(2,10).forEach((r,i)=>console.log(String(i+3).padStart(3)+' | '+r.join(' | ')));
})();
"
```
Expected: 印出改寫後的前幾列，可看到 9 月 1 日為星期二、5 日與 6 日的檢查欄為 `-`

- [ ] **Step 6: 提交**

輸出檢視檔僅供人工確認，不進版控。

```bash
rm -f samples/輸出檢視-115年09月.xlsx
git add src/excel-adapter.js tests/excel-adapter.test.js
git commit -m "新增 Excel 寫入：套用寫入計畫並改寫年月

清空儲存格時保留 s 樣式屬性，年月直接改共用字串內容，
字串數量不變因此 count 與 uniqueCount 無需更動。
測試確認 styles.xml 逐字未變，框線完整保留。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---
### Task 11: Word 讀取與表格解析

把 `.docx` 的表格解成二維字串表格。同樣以正規式處理，不使用 DOM 解析器。

**Files:**
- Create: `src/word-adapter.js`
- Create: `tests/word-adapter.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: `XmlText`、JSZip（參數注入）
- Produces:
  - `Segment = {raw, openTag, text, start, end}` — `start`／`end` 為該段文字在所屬段落合併文字中的位置
  - `WordAdapter.splitSegments(container: string): Segment[]` — 取出容器內所有 `<w:t>`，並算出各自的位置
  - `WordAdapter.mergedText(container: string): string` — 容器內所有 `<w:t>` 串接後的文字
  - `WordAdapter.parseTables(documentXml: string): Array<{raw, rows: Array<{raw, cells: Array<{raw, text}>}>}>`
  - `WordAdapter.tableToGrid(table): string[][]`
  - `WordAdapter.loadDocument(data, JSZipLib): Promise<{zip, documentXml, tables, grids}>`
  - 瀏覽器全域變數名：`WordAdapter`

**實測得知的關鍵事實**：Word 會把一段文字切成多個 run。範例檔的標題就是 `<w:r><w:t>機房狀態月檢查表  </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>115年8月</w:t></w:r>`。因此**任何跨文字的偵測都必須先把整段的 `<w:t>` 串接起來**，不能逐個 run 各自判斷。

**巢狀表格**：本工具不支援表格中還有表格的情形。`parseTables` 以非貪婪比對取最外層，若遇到巢狀結構結果會不正確——這種文件已超出設計範圍，屬於可接受的限制。

- [ ] **Step 1: 寫失敗測試 `tests/word-adapter.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const D = isNode ? require('../src/word-adapter.js') : self.WordAdapter;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('Word › 文字片段', function () {
    const 段落 = '<w:p><w:r><w:t xml:space="preserve">機房檢查表  </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>115年8月</w:t></w:r></w:p>';

    it('取得段落合併後的完整文字', function () {
      assertEqual(D.mergedText(段落), '機房檢查表  115年8月');
    });

    it('切出兩個片段', function () {
      assertEqual(D.splitSegments(段落).length, 2);
    });

    it('每個片段記錄其在合併文字中的位置', function () {
      const s = D.splitSegments(段落);
      assertEqual(s[0].start, 0);
      assertEqual(s[0].end, 7);
      assertEqual(s[1].start, 7);
      assertEqual(s[1].end, 13);
    });

    it('片段保留原始開始標籤', function () {
      const s = D.splitSegments(段落);
      assertEqual(s[0].openTag, '<w:t xml:space="preserve">');
      assertEqual(s[1].openTag, '<w:t>');
    });

    it('還原跳脫字元', function () {
      assertEqual(D.mergedText('<w:t>A&amp;B</w:t>'), 'A&B');
    });

    it('沒有文字時回傳空字串與空陣列', function () {
      assertEqual(D.mergedText('<w:p></w:p>'), '');
      assertDeepEqual(D.splitSegments('<w:p></w:p>'), []);
    });
  });

  describe('Word › 表格解析', function () {
    const 文件 =
      '<w:body><w:p><w:r><w:t>標題</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr/>' +
      '<w:tr><w:tc><w:p><w:r><w:t>日期</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>星期</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p></w:p></w:tc></w:tr>' +
      '</w:tbl></w:body>';

    it('找得到一個表格', function () {
      assertEqual(D.parseTables(文件).length, 1);
    });

    it('表格有兩列', function () {
      assertEqual(D.parseTables(文件)[0].rows.length, 2);
    });

    it('每列有兩格', function () {
      assertEqual(D.parseTables(文件)[0].rows[0].cells.length, 2);
    });

    it('讀得出儲存格文字', function () {
      const t = D.parseTables(文件)[0];
      assertEqual(t.rows[0].cells[0].text, '日期');
      assertEqual(t.rows[0].cells[1].text, '星期');
      assertEqual(t.rows[1].cells[0].text, '1');
    });

    it('空儲存格為空字串', function () {
      assertEqual(D.parseTables(文件)[0].rows[1].cells[1].text, '');
    });

    it('轉成二維陣列', function () {
      const grid = D.tableToGrid(D.parseTables(文件)[0]);
      assertDeepEqual(grid, [['日期', '星期'], ['1', '']]);
    });

    it('跨多個 run 的儲存格文字會被串接', function () {
      const doc = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>115</w:t></w:r>' +
        '<w:r><w:t>年8月</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
      assertEqual(D.parseTables(doc)[0].rows[0].cells[0].text, '115年8月');
    });

    it('沒有表格時回傳空陣列', function () {
      assertDeepEqual(D.parseTables('<w:body><w:p/></w:body>'), []);
    });
  });

  describe('Word › 讀取真實範例檔', function () {
    if (!isNode) return;
    const fs = require('fs');
    const JSZip = require('jszip');
    const A = require('../src/table-analysis.js');

    it('讀得到一個表格與 32 列', async function () {
      const doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      assertEqual(doc.tables.length, 1);
      assertEqual(doc.grids[0].length, 32, '表頭一列加 31 個資料列');
    });

    it('欄位偵測在真實檔案上成功', async function () {
      const doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      const cols = A.detectColumns(doc.grids[0]);
      assertEqual(cols.headerRow, 0);
      assertEqual(cols.dateCol, 0);
      assertEqual(cols.weekCol, 1);
      assertEqual(cols.checkCols.length, 5);
    });

    it('備註欄的文字讀得到', async function () {
      const doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      assertEqual(doc.grids[0][3][6], '空調異音已通報維護廠商', '第 3 天的備註');
    });
  });
})();
```

- [ ] **Step 2: 接上進入點並確認測試失敗**

`run-tests.js` 加入 `'./tests/word-adapter.test.js'`；`tests.html` 加入 `src/word-adapter.js` 與該測試檔。

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/word-adapter.js`

- [ ] **Step 3: 寫實作 `src/word-adapter.js`**

```javascript
// Word 轉接層：以正規式直接處理 docx 內部 XML
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'));
  } else {
    root.WordAdapter = factory(root.XmlText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText) {

  function textRegex() {
    return /(<w:t\b[^>]*>)([\s\S]*?)<\/w:t>/g;
  }

  // 取出容器內所有 w:t，並算出各自在合併文字中的起訖位置
  function splitSegments(container) {
    const re = textRegex();
    const out = [];
    let m;
    let cursor = 0;
    while ((m = re.exec(container)) !== null) {
      const text = XmlText.unescape(m[2]);
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
    const tables = [];
    const tblRe = /<w:tbl>([\s\S]*?)<\/w:tbl>/g;
    let tm;
    while ((tm = tblRe.exec(documentXml)) !== null) {
      const tblInner = tm[1];
      const rows = [];

      const trRe = /<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g;
      let rm;
      while ((rm = trRe.exec(tblInner)) !== null) {
        const trInner = rm[1];
        const cells = [];

        const tcRe = /<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g;
        let cm;
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
      const file = zip.file('word/document.xml');
      if (!file) throw new Error('這個檔案裡找不到 Word 內文，可能不是有效的 Word 檔');
      return file.async('string').then(function (documentXml) {
        const tables = parseTables(documentXml);
        return {
          zip: zip,
          documentXml: documentXml,
          tables: tables,
          grids: tables.map(tableToGrid)
        };
      });
    });
  }

  return {
    splitSegments: splitSegments,
    mergedText: mergedText,
    parseTables: parseTables,
    tableToGrid: tableToGrid,
    loadDocument: loadDocument
  };
});
```

- [ ] **Step 4: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 18 項全數通過

- [ ] **Step 5: 提交**

```bash
git add src/word-adapter.js tests/word-adapter.test.js run-tests.js tests.html
git commit -m "新增 Word 讀取與表格解析

Word 會把一段文字切成多個 run，因此任何跨文字的判斷
都先把段落內的 w:t 串接後再處理。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 12: Word 寫入

把寫入計畫套回 docx，並改寫年月。難點在於一段文字被切成多個 run 時的改寫。

**Files:**
- Modify: `src/word-adapter.js`
- Modify: `tests/word-adapter.test.js`

**Interfaces:**
- Consumes: `WritePlan.CellWrite`、`YearMonthText`
- Produces（皆掛在 `WordAdapter` 上）：
  - `replaceRange(container, segments, from, to, replacement): string` — 改寫合併文字中 `[from, to)` 這一段
  - `setCellText(cellXml, value): string`
  - `applyWritesToDocument(documentXml, table, writes): string`
  - `findParagraphYearMonth(documentXml): Array<{paraIndex, raw, text, matches}>`
  - `rewriteParagraphYearMonth(documentXml, targets, targetYear, targetMonth, selectedKeys): string`
  - `saveDocument(zip, documentXml, JSZipLib, outputType): Promise<Blob|Buffer>`

**跨 run 改寫的規則**：一個匹配可能橫跨多個 `<w:t>`。做法是把新文字整段放進**第一個**涉及的片段，其餘涉及的片段只留下匹配範圍以外的部分。

這會使匹配文字統一採用第一個片段的格式。範例檔的標題正是這種情形——「機房狀態月檢查表  」為一般字體、「115年8月」為粗體，但因為年月**完整落在第二個片段內**，改寫只動到那一個片段，粗體得以保留。**只有匹配真的跨片段時才會有格式合併，屬於可接受的取捨。**

- [ ] **Step 1: 在 `tests/word-adapter.test.js` 末尾追加寫入測試**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const D = isNode ? require('../src/word-adapter.js') : self.WordAdapter;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  describe('Word › 範圍改寫', function () {
    it('改寫落在單一片段內的範圍', function () {
      const c = '<w:r><w:t>機房檢查表</w:t></w:r>';
      const out = D.replaceRange(c, D.splitSegments(c), 2, 4, '狀態');
      assertEqual(D.mergedText(out), '機房狀態表');
    });

    it('改寫時保留該片段的開始標籤', function () {
      const c = '<w:r><w:t xml:space="preserve">機房檢查表</w:t></w:r>';
      const out = D.replaceRange(c, D.splitSegments(c), 2, 4, '狀態');
      assertEqual(out.indexOf('xml:space="preserve"') !== -1, true);
    });

    it('改寫跨兩個片段的範圍', function () {
      const c = '<w:r><w:t>115</w:t></w:r><w:r><w:t>年8月</w:t></w:r>';
      const out = D.replaceRange(c, D.splitSegments(c), 0, 6, '115年9月');
      assertEqual(D.mergedText(out), '115年9月');
    });

    it('跨片段改寫後第二個片段只剩範圍外的文字', function () {
      const c = '<w:r><w:t>甲115</w:t></w:r><w:r><w:t>年8月乙</w:t></w:r>';
      const out = D.replaceRange(c, D.splitSegments(c), 1, 7, '115年9月');
      assertEqual(D.mergedText(out), '甲115年9月乙');
    });

    it('新文字中的特殊字元會被跳脫', function () {
      const c = '<w:r><w:t>abc</w:t></w:r>';
      const out = D.replaceRange(c, D.splitSegments(c), 0, 3, 'a<b');
      assertEqual(out.indexOf('&lt;') !== -1, true);
      assertEqual(D.mergedText(out), 'a<b');
    });
  });

  describe('Word › 儲存格寫入', function () {
    it('改寫既有文字', function () {
      const tc = '<w:tc><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '新')), '新');
    });

    it('多個 run 時只留下新值', function () {
      const tc = '<w:tc><w:p><w:r><w:t>甲</w:t></w:r><w:r><w:t>乙</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '丙')), '丙');
    });

    it('清空為空字串', function () {
      const tc = '<w:tc><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      assertEqual(D.mergedText(D.setCellText(tc, '')), '');
    });

    it('原本沒有文字的儲存格能寫入新值', function () {
      const tc = '<w:tc><w:p></w:p></w:tc>';
      const out = D.setCellText(tc, '六');
      assertEqual(D.mergedText(out), '六');
      assertEqual(out.indexOf('<w:tc>') === 0, true, '仍是合法的儲存格');
    });

    it('儲存格屬性完整保留', function () {
      const tc = '<w:tc><w:tcPr><w:tcW w:w="500"/></w:tcPr><w:p><w:r><w:t>舊</w:t></w:r></w:p></w:tc>';
      const out = D.setCellText(tc, '新');
      assertEqual(out.indexOf('<w:tcW w:w="500"/>') !== -1, true);
    });
  });

  describe('Word › 段落年月改寫', function () {
    const doc = '<w:body><w:p><w:r><w:t>機房檢查表  </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>115年8月</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>製表：115年8月</w:t></w:r></w:p></w:body>';

    it('找得到兩個含年月的段落', function () {
      assertEqual(D.findParagraphYearMonth(doc).length, 2);
    });

    it('全部改寫後兩處都變成新年月', function () {
      const targets = D.findParagraphYearMonth(doc);
      const out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, null);
      assertEqual(out.indexOf('115年8月') === -1, true);
      assertEqual((out.match(/115年9月/g) || []).length, 2);
    });

    it('年月完整落在單一 run 時粗體格式得以保留', function () {
      const targets = D.findParagraphYearMonth(doc);
      const out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, null);
      assertEqual(out.indexOf('<w:rPr><w:b/></w:rPr><w:t>115年9月</w:t>') !== -1, true);
    });

    it('只改寫選定的段落', function () {
      const targets = D.findParagraphYearMonth(doc);
      const out = D.rewriteParagraphYearMonth(doc, targets, 2026, 9, ['0-0']);
      assertEqual(out.indexOf('製表：115年8月') !== -1, true, '未選取的維持原狀');
    });

    it('年月被切成多個 run 時仍改寫得到', function () {
      const d2 = '<w:body><w:p><w:r><w:t>115</w:t></w:r><w:r><w:t>年8月報表</w:t></w:r></w:p></w:body>';
      const targets = D.findParagraphYearMonth(d2);
      assertEqual(targets.length, 1);
      const out = D.rewriteParagraphYearMonth(d2, targets, 2026, 9, null);
      assertEqual(D.mergedText(out), '115年9月報表');
    });

    it('表格內的年月不會被段落改寫重複處理', function () {
      const d3 = '<w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>115年8月</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>';
      const targets = D.findParagraphYearMonth(d3);
      assertEqual(targets.length, 1, '表格內的段落也算一個目標');
    });
  });

  describe('Word › 完整改寫真實範例檔', function () {
    if (!isNode) return;
    const fs = require('fs');
    const JSZip = require('jszip');
    const A = require('../src/table-analysis.js');
    const M = require('../src/month-model.js');
    const H = require('../src/holidays-tw.js');
    const B = require('../src/holidays-builtin.js');
    const W = require('../src/write-plan.js');

    async function 產生改寫結果() {
      const doc = await D.loadDocument(fs.readFileSync('samples/機房月檢查表-範例.docx'), JSZip);
      const grid = doc.grids[0];
      const cols = A.detectColumns(grid);
      const model = M.buildMonth(2026, 9, H.normalizeYearData(B['2026']));
      const plan = W.buildWritePlan({
        monthModel: model, table: grid, headerRow: cols.headerRow,
        dateCol: cols.dateCol, weekCol: cols.weekCol, checkCols: cols.checkCols
      });
      let xml = D.applyWritesToDocument(doc.documentXml, doc.tables[0], plan.writes);
      const targets = D.findParagraphYearMonth(xml);
      xml = D.rewriteParagraphYearMonth(xml, targets, 2026, 9, null);
      const out = await D.saveDocument(doc.zip, xml, JSZip, 'nodebuffer');
      return { out: out, plan: plan };
    }

    it('改寫後的檔案仍可被重新讀取', async function () {
      const r = await 產生改寫結果();
      const doc2 = await D.loadDocument(r.out, JSZip);
      assertEqual(doc2.grids[0][0][0], '日期', '表頭未受影響');
    });

    it('日期與星期正確寫入', async function () {
      const r = await 產生改寫結果();
      const g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[1][0], '1');
      assertEqual(g[1][1], '二', '9月1日為星期二');
      assertEqual(g[5][0], '5');
      assertEqual(g[5][1], '六', '9月5日為星期六');
    });

    it('假日的檢查欄填入減號', async function () {
      const r = await 產生改寫結果();
      const g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[5][2], '-', '9月5日為週六');
      assertEqual(g[25][2], '-', '9月25日為中秋節');
    });

    it('平日的檢查欄被清空', async function () {
      const r = await 產生改寫結果();
      const g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[1][2], '', '9月1日為平日');
    });

    it('超出當月天數的第 31 列被整列清空', async function () {
      const r = await 產生改寫結果();
      const g = (await D.loadDocument(r.out, JSZip)).grids[0];
      assertEqual(g[31][0], '', '9月只有 30 天');
      assertEqual(g[31][1], '');
    });

    it('標題的年月被改寫且保留粗體', async function () {
      const r = await 產生改寫結果();
      const doc2 = await D.loadDocument(r.out, JSZip);
      assertEqual(doc2.documentXml.indexOf('115年9月') !== -1, true);
      assertEqual(doc2.documentXml.indexOf('115年8月') === -1, true);
      assertEqual(doc2.documentXml.indexOf('<w:b/></w:rPr><w:t>115年9月</w:t>') !== -1, true,
        '年月完整落在粗體 run 內，粗體應保留');
    });

    it('表格框線設定完全未被更動', async function () {
      const r = await 產生改寫結果();
      const 新檔 = await D.loadDocument(r.out, JSZip);
      assertEqual(新檔.documentXml.indexOf('<w:tblBorders>') !== -1, true);
      assertEqual(新檔.documentXml.indexOf('<w:insideH w:val="single"') !== -1, true);
    });
  });
})();
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `node run-tests.js`
Expected: FAIL，`D.replaceRange is not a function`

- [ ] **Step 3: 在 `src/word-adapter.js` 新增寫入函式**

模組頭尾改為同時注入 `YearMonthText`：

```javascript
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xml-text.js'), require('./year-month-text.js'));
  } else {
    root.WordAdapter = factory(root.XmlText, root.YearMonthText);
  }
})(typeof self !== 'undefined' ? self : this, function (XmlText, YearMonthText) {
```

新增的函式：

```javascript
  // 改寫合併文字中 [from, to) 這一段。
  // 新文字整段放進第一個涉及的片段，其餘涉及的片段只留範圍外的部分。
  function replaceRange(container, segments, from, to, replacement) {
    const touched = segments.filter(function (s) {
      return s.start < to && s.end > from;
    });
    if (touched.length === 0) return container;

    const updates = [];
    touched.forEach(function (seg, i) {
      const 前段 = seg.text.slice(0, Math.max(0, from - seg.start));
      const 後段 = seg.text.slice(Math.max(0, to - seg.start));
      const 新文字 = i === 0 ? 前段 + replacement + 後段 : 前段 + 後段;
      updates.push({
        raw: seg.raw,
        replacement: seg.openTag + XmlText.escape(新文字) + '</w:t>'
      });
    });

    // 由後往前替換，避免前面的長度變化影響後面的位置
    let out = container;
    for (let i = updates.length - 1; i >= 0; i--) {
      const at = out.lastIndexOf(updates[i].raw);
      if (at === -1) continue;
      out = out.slice(0, at) + updates[i].replacement + out.slice(at + updates[i].raw.length);
    }
    return out;
  }

  function setCellText(cellXml, value) {
    const segments = splitSegments(cellXml);

    if (segments.length > 0) {
      const total = segments[segments.length - 1].end;
      return replaceRange(cellXml, segments, 0, total, value);
    }

    // 原本沒有任何文字，需要插入一個 run
    if (value === '') return cellXml;
    const newRun = '<w:r><w:t xml:space="preserve">' + XmlText.escape(value) + '</w:t></w:r>';
    const at = cellXml.lastIndexOf('</w:p>');
    if (at === -1) return cellXml;
    return cellXml.slice(0, at) + newRun + cellXml.slice(at);
  }

  function applyWritesToDocument(documentXml, table, writes) {
    // 依列分組，一次重建一整列，避免逐格替換時位置錯亂
    const byRow = {};
    writes.forEach(function (w) {
      if (!byRow[w.row]) byRow[w.row] = [];
      byRow[w.row].push(w);
    });

    let newTableXml = table.raw;

    Object.keys(byRow).forEach(function (rowKey) {
      const rowIndex = parseInt(rowKey, 10);
      const row = table.rows[rowIndex];
      if (!row) return;

      let newRowXml = row.raw;
      byRow[rowKey].forEach(function (w) {
        const cell = row.cells[w.col];
        if (!cell) return;
        const newCellXml = setCellText(cell.raw, w.value);
        const at = newRowXml.indexOf(cell.raw);
        if (at === -1) return;
        newRowXml = newRowXml.slice(0, at) + newCellXml + newRowXml.slice(at + cell.raw.length);
      });

      const at = newTableXml.indexOf(row.raw);
      if (at === -1) return;
      newTableXml = newTableXml.slice(0, at) + newRowXml + newTableXml.slice(at + row.raw.length);
    });

    const at = documentXml.indexOf(table.raw);
    if (at === -1) return documentXml;
    return documentXml.slice(0, at) + newTableXml + documentXml.slice(at + table.raw.length);
  }

  function findParagraphYearMonth(documentXml) {
    const out = [];
    const paraRe = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
    let m;
    let index = 0;
    while ((m = paraRe.exec(documentXml)) !== null) {
      const raw = m[0];
      const text = mergedText(raw);
      const matches = YearMonthText.findYearMonth(text);
      if (matches.length > 0) {
        out.push({ paraIndex: index, raw: raw, text: text, matches: matches });
      }
      index++;
    }
    return out;
  }

  function rewriteParagraphYearMonth(documentXml, targets, targetYear, targetMonth, selectedKeys) {
    let out = documentXml;

    targets.forEach(function (target) {
      let newParaXml = target.raw;
      // 由後往前處理各個匹配，避免長度變化影響前面的位置
      const ordered = target.matches
        .map(function (m, i) { return { match: m, i: i }; })
        .sort(function (a, b) { return b.match.index - a.match.index; });

      let 有改動 = false;
      ordered.forEach(function (entry) {
        const key = target.paraIndex + '-' + entry.i;
        if (selectedKeys && selectedKeys.indexOf(key) === -1) return;
        const 新文字 = YearMonthText.rewriteMatch(entry.match, targetYear, targetMonth);
        if (新文字 === entry.match.raw) return;
        newParaXml = replaceRange(
          newParaXml, splitSegments(newParaXml),
          entry.match.index, entry.match.index + entry.match.raw.length,
          新文字
        );
        有改動 = true;
      });

      if (!有改動) return;
      const at = out.indexOf(target.raw);
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
```

把這六個函式加進 `return` 物件。

- [ ] **Step 4: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 23 項全數通過

- [ ] **Step 5: 提交**

```bash
git add src/word-adapter.js tests/word-adapter.test.js
git commit -m "新增 Word 寫入：套用寫入計畫並改寫年月

支援跨多個 run 的文字改寫；年月完整落在單一 run 時
只動該 run，字體格式得以保留。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---
### Task 13: 介面骨架與月曆

建立開發用的介面進入點，包含年月選擇與當月月曆。**年月預設取自使用者電腦的系統時間**，不需手動指定；選擇器是為了補做上個月或預先做下個月而存在。

**Files:**
- Create: `src/ui-calendar.js`
- Create: `src/app.css`
- Create: `src/index.html`
- Create: `tests/ui-calendar.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: `MonthModel.DayModel`、`CalendarCore.WEEKDAY_NAMES`
- Produces:
  - `UiCalendar.defaultYearMonth(now?: Date): {year, month}` — 省略參數時取當下系統時間
  - `UiCalendar.buildCalendarWeeks(monthModel, firstWeekdayIndex): Array<Array<DayModel|null>>`
  - `UiCalendar.weekdayIndex(weekday: string): number` — `'日'` → 0
  - `UiCalendar.renderCalendar(monthModel, container, onToggle): void`
  - `UiCalendar.sourceLabel(source: string, year: number): string`
  - 瀏覽器全域變數名：`UiCalendar`

- [ ] **Step 1: 寫失敗測試 `tests/ui-calendar.test.js`**

只測純函式，畫面渲染留待人工確認。

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const U = isNode ? require('../src/ui-calendar.js') : self.UiCalendar;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual;

  describe('介面 › 預設年月', function () {
    it('取自傳入的日期', function () {
      const r = U.defaultYearMonth(new Date(2026, 8, 30));
      assertEqual(r.year, 2026);
      assertEqual(r.month, 9, '月份為 1 起算');
    });

    it('一月不會變成零月', function () {
      assertEqual(U.defaultYearMonth(new Date(2026, 0, 15)).month, 1);
    });

    it('十二月正確', function () {
      assertEqual(U.defaultYearMonth(new Date(2026, 11, 1)).month, 12);
    });

    it('省略參數時取系統時間且落在合理範圍', function () {
      const r = U.defaultYearMonth();
      assertEqual(r.month >= 1 && r.month <= 12, true);
      assertEqual(r.year >= 2020 && r.year <= 2100, true);
    });
  });

  describe('介面 › 星期索引', function () {
    it('日為 0、六為 6', function () {
      assertEqual(U.weekdayIndex('日'), 0);
      assertEqual(U.weekdayIndex('一'), 1);
      assertEqual(U.weekdayIndex('六'), 6);
    });

    it('無法辨識時回傳 -1', function () {
      assertEqual(U.weekdayIndex('X'), -1);
    });
  });

  describe('介面 › 月曆排列', function () {
    // 2026 年 9 月：1 日為星期二
    const 模型 = [];
    const 星期 = ['二','三','四','五','六','日','一'];
    for (let d = 1; d <= 30; d++) {
      模型.push({ day: d, weekday: 星期[(d - 1) % 7], isHoliday: false, isMakeup: false, description: '' });
    }

    it('第一週開頭補上空格', function () {
      const weeks = U.buildCalendarWeeks(模型, 2);
      assertEqual(weeks[0][0], null, '星期日位置為空');
      assertEqual(weeks[0][1], null, '星期一位置為空');
      assertEqual(weeks[0][2].day, 1, '星期二為 1 日');
    });

    it('每一週都是七格', function () {
      U.buildCalendarWeeks(模型, 2).forEach(function (w) { assertEqual(w.length, 7); });
    });

    it('最後一天位置正確', function () {
      const weeks = U.buildCalendarWeeks(模型, 2);
      const last = weeks[weeks.length - 1];
      const 有值 = last.filter(function (c) { return c !== null; });
      assertEqual(有值[有值.length - 1].day, 30);
    });

    it('月初為星期日時第一週不補空格', function () {
      const m = [{ day: 1, weekday: '日', isHoliday: true, isMakeup: false, description: '' }];
      assertEqual(U.buildCalendarWeeks(m, 0)[0][0].day, 1);
    });
  });

  describe('介面 › 資料來源標示', function () {
    it('線上取得的說明文字', function () {
      assertEqual(U.sourceLabel('online', 2026), '已取得 2026 年最新假日資料');
    });

    it('內建資料的說明文字', function () {
      assertEqual(U.sourceLabel('builtin', 2026), '使用內建的 2026 年假日資料（無法連線）');
    });

    it('純計算的說明文字要帶警告', function () {
      assertEqual(U.sourceLabel('computed', 2031), '⚠ 沒有 2031 年的假日資料，僅依週六日判斷');
    });
  });
})();
```

- [ ] **Step 2: 接上進入點並確認測試失敗**

`run-tests.js` 加入 `'./tests/ui-calendar.test.js'`；`tests.html` 加入 `src/ui-calendar.js` 與該測試檔。

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/ui-calendar.js`

- [ ] **Step 3: 寫實作 `src/ui-calendar.js`**

```javascript
// 介面：年月選擇與月曆呈現
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.UiCalendar = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const WEEK_ORDER = ['日', '一', '二', '三', '四', '五', '六'];

  // 預設取使用者電腦的系統時間，不需手動指定
  function defaultYearMonth(now) {
    const d = now || new Date();
    return { year: d.getFullYear(), month: d.getMonth() + 1 };
  }

  function weekdayIndex(weekday) {
    return WEEK_ORDER.indexOf(weekday);
  }

  function buildCalendarWeeks(monthModel, firstWeekdayIndex) {
    const weeks = [];
    let week = [];
    for (let i = 0; i < firstWeekdayIndex; i++) week.push(null);

    monthModel.forEach(function (day) {
      week.push(day);
      if (week.length === 7) { weeks.push(week); week = []; }
    });

    if (week.length > 0) {
      while (week.length < 7) week.push(null);
      weeks.push(week);
    }
    return weeks;
  }

  function sourceLabel(source, year) {
    if (source === 'online') return '已取得 ' + year + ' 年最新假日資料';
    if (source === 'builtin') return '使用內建的 ' + year + ' 年假日資料（無法連線）';
    return '⚠ 沒有 ' + year + ' 年的假日資料，僅依週六日判斷';
  }

  // 渲染月曆。點擊某一天會呼叫 onToggle(day)，供使用者修正颱風假等臨時狀況。
  function renderCalendar(monthModel, container, onToggle) {
    container.innerHTML = '';
    if (monthModel.length === 0) return;

    const first = weekdayIndex(monthModel[0].weekday);
    const weeks = buildCalendarWeeks(monthModel, first < 0 ? 0 : first);

    const table = document.createElement('table');
    table.className = 'calendar';

    const thead = document.createElement('thead');
    const hrow = document.createElement('tr');
    WEEK_ORDER.forEach(function (w) {
      const th = document.createElement('th');
      th.textContent = w;
      hrow.appendChild(th);
    });
    thead.appendChild(hrow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    weeks.forEach(function (week) {
      const tr = document.createElement('tr');
      week.forEach(function (day) {
        const td = document.createElement('td');
        if (day) {
          td.textContent = day.day;
          td.title = day.description || '';
          if (day.isHoliday) td.className = 'holiday';
          else if (day.isMakeup) td.className = 'makeup';
          if (day.description) td.classList.add('named');
          td.addEventListener('click', function () { onToggle(day); });
        } else {
          td.className = 'blank';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    container.appendChild(table);
  }

  return {
    WEEK_ORDER: WEEK_ORDER,
    defaultYearMonth: defaultYearMonth,
    weekdayIndex: weekdayIndex,
    buildCalendarWeeks: buildCalendarWeeks,
    sourceLabel: sourceLabel,
    renderCalendar: renderCalendar
  };
});
```

- [ ] **Step 4: 寫樣式 `src/app.css`**

```css
:root {
  --bg: #f7f7f5;
  --panel: #ffffff;
  --line: #dcdcd6;
  --text: #23231f;
  --muted: #6b6b63;
  --holiday: #b3261e;
  --holiday-bg: #fdecea;
  --makeup: #1a4f8a;
  --makeup-bg: #e8f0fa;
  --accent: #2d5a3d;
}

* { box-sizing: border-box; }

body {
  font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
  margin: 0;
  padding: 1.5rem;
  background: var(--bg);
  color: var(--text);
  line-height: 1.6;
}

h1 { font-size: 1.3rem; margin: 0 0 1rem; }
h2 { font-size: 1rem; margin: 0 0 .6rem; color: var(--muted); font-weight: 600; }

.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 1rem 1.2rem;
  margin-bottom: 1rem;
}

.row { display: flex; gap: .8rem; align-items: center; flex-wrap: wrap; }

select, button {
  font: inherit;
  padding: .4rem .8rem;
  border: 1px solid var(--line);
  border-radius: 5px;
  background: #fff;
}

button.primary {
  background: var(--accent);
  color: #fff;
  border-color: var(--accent);
  cursor: pointer;
}
button.primary:disabled { opacity: .45; cursor: not-allowed; }

#drop {
  border: 2px dashed var(--line);
  border-radius: 8px;
  padding: 2rem;
  text-align: center;
  color: var(--muted);
  cursor: pointer;
}
#drop.over { border-color: var(--accent); background: #f0f5f1; }

table.calendar { border-collapse: collapse; }
table.calendar th, table.calendar td {
  width: 2.6rem; height: 2.3rem;
  text-align: center;
  border: 1px solid var(--line);
  font-size: .9rem;
}
table.calendar th { background: #f0f0ec; font-weight: 600; }
table.calendar td { cursor: pointer; }
table.calendar td.blank { background: #fafaf8; cursor: default; }
table.calendar td.holiday { color: var(--holiday); background: var(--holiday-bg); font-weight: 600; }
table.calendar td.makeup { color: var(--makeup); background: var(--makeup-bg); font-weight: 600; }
table.calendar td.named::after { content: "•"; display: block; font-size: .6rem; line-height: 0; }

table.preview { border-collapse: collapse; font-size: .85rem; width: 100%; }
table.preview th, table.preview td {
  border: 1px solid var(--line);
  padding: .25rem .5rem;
  text-align: center;
}
table.preview th { background: #f0f0ec; }
tr.is-holiday td { background: var(--holiday-bg); }
tr.is-makeup td { background: var(--makeup-bg); }

.warn {
  background: #fff8e1;
  border: 1px solid #e6cf7a;
  border-radius: 6px;
  padding: .7rem .9rem;
  margin: .6rem 0;
  font-size: .9rem;
}
.warn strong { color: #8a6100; }

.error {
  background: var(--holiday-bg);
  border: 1px solid #e8a49e;
  border-radius: 6px;
  padding: .7rem .9rem;
  color: var(--holiday);
}

.muted { color: var(--muted); font-size: .85rem; }
.scroll { max-height: 22rem; overflow: auto; }
.hidden { display: none; }
```

- [ ] **Step 5: 寫開發用進入點 `src/index.html`**

此檔載入所有模組，供開發期間直接雙擊開啟測試。`build.js` 稍後會依它合成單一發布檔。

```html
<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>機房月檢查表 星期填寫工具</title>
<link rel="stylesheet" href="app.css">
</head>
<body>
<h1>機房月檢查表　星期填寫工具</h1>

<div class="panel">
  <h2>一、選擇年月</h2>
  <div class="row">
    <select id="year"></select>
    <select id="month"></select>
    <span id="source" class="muted"></span>
  </div>
  <div id="calendar" style="margin-top:.8rem"></div>
  <p class="muted">紅色為假日、藍色為補班日。點一下可以切換某天的假日狀態（例如颱風假）。</p>
</div>

<div class="panel">
  <h2>二、選擇檔案</h2>
  <div id="drop">把 .xlsx 或 .docx 檔拖到這裡，或點一下選擇檔案</div>
  <input type="file" id="file" accept=".xlsx,.docx" class="hidden">
  <div id="fileInfo" class="muted" style="margin-top:.6rem"></div>
</div>

<div class="panel hidden" id="manualPanel">
  <h2>指定欄位</h2>
  <p class="muted">工具找不到「日期」與「星期」欄位標題，請直接指定是哪幾欄。指定過一次之後，
    下次遇到同樣格式的表格就會自動帶入。</p>
  <div class="scroll"><div id="manualTable"></div></div>
  <div class="row" style="margin-top:.8rem">
    <label>表頭在第幾列 <select id="manualHeader"></select></label>
    <label>日期欄 <select id="manualDate"></select></label>
    <label>星期欄 <select id="manualWeek"></select></label>
    <button id="manualConfirm" class="primary">確定</button>
  </div>
</div>

<div class="panel hidden" id="previewPanel">
  <h2>三、確認後下載</h2>
  <div id="messages"></div>
  <div id="yearMonthList"></div>
  <div class="scroll"><div id="preview"></div></div>
  <div class="row" style="margin-top:.9rem">
    <button id="download" class="primary">產生檔案並下載</button>
    <span id="outputName" class="muted"></span>
  </div>
</div>

<script src="../vendor/jszip.min.js"></script>
<script src="calendar-core.js"></script>
<script src="holidays-builtin.js"></script>
<script src="holidays-tw.js"></script>
<script src="month-model.js"></script>
<script src="xml-text.js"></script>
<script src="year-month-text.js"></script>
<script src="table-analysis.js"></script>
<script src="write-plan.js"></script>
<script src="excel-adapter.js"></script>
<script src="word-adapter.js"></script>
<script src="ui-calendar.js"></script>
<script src="ui-main.js"></script>
</body>
</html>
```

- [ ] **Step 6: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 13 項全數通過

- [ ] **Step 7: 提交**

`ui-main.js` 於下一個任務建立，此時開啟 `index.html` 會有一個載入失敗的錯誤，屬正常。

```bash
git add src/ui-calendar.js src/app.css src/index.html tests/ui-calendar.test.js run-tests.js tests.html
git commit -m "新增介面骨架與月曆呈現

年月預設取自系統時間，月曆標示假日與補班日並可點擊修正。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 14: 介面主流程

串起完整流程：選年月 → 取假日 → 拖入檔案 → 偵測欄位 → 預覽 → 下載。

**Files:**
- Create: `src/ui-main.js`
- Create: `tests/ui-main.test.js`
- Modify: `run-tests.js`、`tests.html`

**Interfaces:**
- Consumes: 前面所有模組
- Produces:
  - `UiMain.outputFilename(originalName: string, year: number, month: number): string`
  - `UiMain.detectFormat(filename: string): 'xlsx'|'docx'|'legacy'|'unknown'`
  - `UiMain.storageKey(headers: string[]): string`
  - `UiMain.columnsFrom(headerCells: string[], headerRow: number, dateCol: number, weekCol: number): {headerRow, dateCol, weekCol, checkCols}`
  - 瀏覽器全域變數名：`UiMain`

**輸出檔名規則**：`原檔名_115年09月.副檔名`，年份採民國、月份補零。

**舊格式處理**：`.xls` 與 `.doc` 明確擋下並提示另存新檔，不可默默失敗。

**跨站指令碼防護**：預覽區與手動指定區都以字串拼接產生 HTML，而其中的表頭、儲存格值、
備註文字全部來自使用者拖入的檔案。這些內容必須先經 `XmlText.escape` 跳脫再放進 `innerHTML`，
否則一個內容為 `<img src=x onerror=alert(1)>` 的儲存格就會在頁面上執行程式碼。
工具雖然在本機執行，但檔案未必出自使用者之手，這道防護不可省略。

**欄位偵測失敗的退路**（spec 第六節）：自動偵測失敗時不可直接報錯了事，必須顯示表格前幾列，
讓使用者自己指定表頭列、日期欄與星期欄。指定結果以表頭文字為鍵存入 `localStorage`，
下次遇到相同格式的表格自動帶入。`localStorage` 在部分瀏覽器設定下會擲出例外，
因此讀寫一律包在 `try/catch` 中，失敗時只是失去記憶功能，不影響主要流程。

- [ ] **Step 1: 寫失敗測試 `tests/ui-main.test.js`**

```javascript
(function () {
  const isNode = typeof module === 'object' && module.exports;
  const T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  const U = isNode ? require('../src/ui-main.js') : self.UiMain;
  const describe = T.describe, it = T.it;
  const assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual;

  describe('介面 › 輸出檔名', function () {
    it('加上民國年月且月份補零', function () {
      assertEqual(U.outputFilename('機房檢查表.xlsx', 2026, 9), '機房檢查表_115年09月.xlsx');
    });

    it('兩位數月份不再補零', function () {
      assertEqual(U.outputFilename('檢查表.docx', 2026, 12), '檢查表_115年12月.docx');
    });

    it('檔名中含有點號時只取最後一段為副檔名', function () {
      assertEqual(U.outputFilename('機房.狀態.檢查表.xlsx', 2026, 9),
        '機房.狀態.檢查表_115年09月.xlsx');
    });

    it('跨年時民國年正確', function () {
      assertEqual(U.outputFilename('表.xlsx', 2027, 1), '表_116年01月.xlsx');
    });

    it('沒有副檔名時直接附加', function () {
      assertEqual(U.outputFilename('檢查表', 2026, 9), '檢查表_115年09月');
    });
  });

  describe('介面 › 格式判斷', function () {
    it('辨識 xlsx 與 docx', function () {
      assertEqual(U.detectFormat('a.xlsx'), 'xlsx');
      assertEqual(U.detectFormat('a.docx'), 'docx');
    });

    it('大寫副檔名也辨識得出', function () {
      assertEqual(U.detectFormat('A.XLSX'), 'xlsx');
    });

    it('舊格式獨立標示以便給出明確提示', function () {
      assertEqual(U.detectFormat('a.xls'), 'legacy');
      assertEqual(U.detectFormat('a.doc'), 'legacy');
    });

    it('其他格式為未知', function () {
      assertEqual(U.detectFormat('a.pdf'), 'unknown');
      assertEqual(U.detectFormat('a'), 'unknown');
    });
  });

  describe('介面 › 欄位記憶鍵值', function () {
    it('相同表頭產生相同鍵值', function () {
      const a = U.storageKey(['日期', '星期', '溫濕度']);
      const b = U.storageKey(['日期', '星期', '溫濕度']);
      assertEqual(a, b);
    });

    it('不同表頭產生不同鍵值', function () {
      const a = U.storageKey(['日期', '星期', '溫濕度']);
      const b = U.storageKey(['日期', '星期', 'UPS']);
      assertEqual(a === b, false);
    });

    it('鍵值有固定前綴便於辨識', function () {
      assertEqual(U.storageKey(['日期']).indexOf('week-automodify:') === 0, true);
    });
  });

  describe('介面 › 由手動指定組出欄位設定', function () {
    it('其餘欄位皆列為檢查欄', function () {
      const r = U.columnsFrom(['日期', '星期', '溫濕度', '備註'], 2, 0, 1);
      assertEqual(r.headerRow, 2);
      assertEqual(r.dateCol, 0);
      assertEqual(r.weekCol, 1);
      assertDeepEqual(r.checkCols, [2, 3]);
    });

    it('日期與星期不在最前面兩欄也正確', function () {
      const r = U.columnsFrom(['項次', '日期', '星期', '溫濕度'], 0, 1, 2);
      assertDeepEqual(r.checkCols, [0, 3]);
    });
  });
})();
```

- [ ] **Step 2: 接上進入點並確認測試失敗**

`run-tests.js` 加入 `'./tests/ui-main.test.js'`；`tests.html` 加入 `src/ui-main.js` 與該測試檔。

Run: `node run-tests.js`
Expected: FAIL，找不到模組 `../src/ui-main.js`

- [ ] **Step 3: 寫實作 `src/ui-main.js`**

模組同時匯出純函式（供測試）與在瀏覽器中啟動介面。

```javascript
// 介面主流程：串起選年月、取假日、讀檔、預覽、下載
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else {
    root.UiMain = api;
    // 只有在真正的工具頁面才啟動介面。tests.html 也會載入本檔，
    // 但沒有這些元素，不可貿然啟動。
    if (typeof document !== 'undefined' && document.getElementById('drop')) api.start();
  }
})(typeof self !== 'undefined' ? self : this, function () {

  const ROC_OFFSET = 1911;

  function pad2(n) { return n < 10 ? '0' + n : String(n); }

  function outputFilename(originalName, year, month) {
    const 後綴 = '_' + (year - ROC_OFFSET) + '年' + pad2(month) + '月';
    const at = originalName.lastIndexOf('.');
    if (at <= 0) return originalName + 後綴;
    return originalName.slice(0, at) + 後綴 + originalName.slice(at);
  }

  function detectFormat(filename) {
    const lower = String(filename).toLowerCase();
    if (/\.xlsx$/.test(lower)) return 'xlsx';
    if (/\.docx$/.test(lower)) return 'docx';
    if (/\.(xls|doc)$/.test(lower)) return 'legacy';
    return 'unknown';
  }

  function storageKey(headers) {
    return 'week-automodify:' + headers.join('|');
  }

  // 由使用者指定的日期欄與星期欄，組出完整的欄位設定
  function columnsFrom(headerCells, headerRow, dateCol, weekCol) {
    const checkCols = [];
    for (let c = 0; c < headerCells.length; c++) {
      if (c !== dateCol && c !== weekCol) checkCols.push(c);
    }
    return { headerRow: headerRow, dateCol: dateCol, weekCol: weekCol, checkCols: checkCols };
  }

  // 以下僅在瀏覽器中執行
  function start() {
    const $ = function (id) { return document.getElementById(id); };
    const state = {
      year: null, month: null, model: [], source: '',
      file: null, format: null, parsed: null, cols: null, grid: null,
      plan: null, yearMonthTargets: []
    };

    // 來自檔案的文字一律跳脫後才放進 HTML。
    // 儲存格內容是使用者拖進來的資料，若直接拼接，
    // 一個含 <img onerror=...> 的儲存格就能在頁面上執行程式碼。
    function esc(v) {
      return XmlText.escape(v === null || v === undefined ? '' : String(v));
    }

    function 顯示錯誤(msg) {
      $('messages').innerHTML = '<div class="error">' + esc(msg) + '</div>';
      $('previewPanel').classList.remove('hidden');
    }

    function 初始化選單() {
      const d = UiCalendar.defaultYearMonth();
      state.year = d.year;
      state.month = d.month;

      const ys = $('year');
      for (let y = d.year - 2; y <= d.year + 2; y++) {
        const o = document.createElement('option');
        o.value = y;
        o.textContent = y + '年（民國 ' + (y - ROC_OFFSET) + ' 年）';
        if (y === d.year) o.selected = true;
        ys.appendChild(o);
      }
      const ms = $('month');
      for (let m = 1; m <= 12; m++) {
        const o = document.createElement('option');
        o.value = m;
        o.textContent = m + ' 月';
        if (m === d.month) o.selected = true;
        ms.appendChild(o);
      }
      ys.addEventListener('change', function () { state.year = parseInt(ys.value, 10); 重新整理月份(); });
      ms.addEventListener('change', function () { state.month = parseInt(ms.value, 10); 重新整理月份(); });
    }

    function 重新整理月份() {
      $('source').textContent = '讀取假日資料中……';
      HolidaysTW.getYearData(state.year, { builtin: HolidaysBuiltin }).then(function (r) {
        state.source = r.source;
        state.model = MonthModel.buildMonth(state.year, state.month, r.data);
        $('source').textContent = UiCalendar.sourceLabel(r.source, state.year);
        UiCalendar.renderCalendar(state.model, $('calendar'), 切換假日);
        if (state.parsed) 重新整理預覽();
      });
    }

    function 切換假日(day) {
      day.isHoliday = !day.isHoliday;
      day.isMakeup = false;
      UiCalendar.renderCalendar(state.model, $('calendar'), 切換假日);
      if (state.parsed) 重新整理預覽();
    }

    function 讀取檔案(file) {
      state.file = file;
      state.format = detectFormat(file.name);
      $('fileInfo').textContent = file.name;

      if (state.format === 'legacy') {
        return 顯示錯誤('這是舊版的 Office 格式。請用 Excel 或 Word 開啟後，' +
          '另存成 .xlsx 或 .docx 再試一次。');
      }
      if (state.format === 'unknown') {
        return 顯示錯誤('只支援 .xlsx 與 .docx 兩種格式。');
      }

      file.arrayBuffer().then(function (buf) {
        const 載入 = state.format === 'xlsx'
          ? ExcelAdapter.loadWorkbook(buf, JSZip)
          : WordAdapter.loadDocument(buf, JSZip);
        return 載入;
      }).then(function (parsed) {
        state.parsed = parsed;
        const grid = state.format === 'xlsx' ? parsed.table : (parsed.grids[0] || []);
        if (!grid || grid.length === 0) {
          return 顯示錯誤('這個檔案裡找不到表格。');
        }
        state.grid = grid;
        state.cols = TableAnalysis.detectColumns(grid) || 從記憶還原欄位(grid);
        if (!state.cols) {
          return 顯示手動指定(grid);
        }
        $('manualPanel').classList.add('hidden');
        重新整理預覽();
      }).catch(function (err) {
        顯示錯誤('讀取失敗：' + err.message);
      });
    }

    function 讀取記憶(key) {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : null;
      } catch (e) {
        return null; // 瀏覽器停用儲存時只是失去記憶功能
      }
    }

    function 寫入記憶(key, value) {
      try {
        localStorage.setItem(key, JSON.stringify(value));
      } catch (e) {
        // 忽略，不影響主要流程
      }
    }

    // 自動偵測失敗時，看看這張表的格式是不是以前指定過
    function 從記憶還原欄位(grid) {
      const limit = Math.min(grid.length, 10);
      for (let r = 0; r < limit; r++) {
        const cells = (grid[r] || []).map(function (x) { return String(x === null || x === undefined ? '' : x); });
        if (cells.length === 0) continue;
        const saved = 讀取記憶(storageKey(cells));
        if (saved) return columnsFrom(cells, r, saved.dateCol, saved.weekCol);
      }
      return null;
    }

    function 顯示手動指定(grid) {
      const limit = Math.min(grid.length, 10);
      const html = ['<table class="preview"><tbody>'];
      for (let r = 0; r < limit; r++) {
        html.push('<tr><th>第 ' + (r + 1) + ' 列</th>');
        (grid[r] || []).forEach(function (c) { html.push('<td>' + esc(c) + '</td>'); });
        html.push('</tr>');
      }
      html.push('</tbody></table>');
      $('manualTable').innerHTML = html.join('');

      const 列選單 = $('manualHeader');
      列選單.innerHTML = '';
      for (let r = 0; r < limit; r++) {
        const o = document.createElement('option');
        o.value = r;
        o.textContent = '第 ' + (r + 1) + ' 列';
        列選單.appendChild(o);
      }
      填入欄選單();
      列選單.onchange = 填入欄選單;
      $('manualPanel').classList.remove('hidden');
      $('previewPanel').classList.add('hidden');
    }

    function 填入欄選單() {
      const r = parseInt($('manualHeader').value, 10);
      const cells = state.grid[r] || [];
      [$('manualDate'), $('manualWeek')].forEach(function (sel) {
        sel.innerHTML = '';
        cells.forEach(function (c, i) {
          const o = document.createElement('option');
          o.value = i;
          o.textContent = '第 ' + (i + 1) + ' 欄' + (c ? '（' + c + '）' : '');
          sel.appendChild(o);
        });
      });
      if (cells.length > 1) $('manualWeek').value = 1;
    }

    function 套用手動指定() {
      const r = parseInt($('manualHeader').value, 10);
      const dateCol = parseInt($('manualDate').value, 10);
      const weekCol = parseInt($('manualWeek').value, 10);
      if (dateCol === weekCol) {
        return 顯示錯誤('日期欄與星期欄不能是同一欄。');
      }
      const cells = (state.grid[r] || []).map(function (x) { return String(x === null || x === undefined ? '' : x); });
      state.cols = columnsFrom(cells, r, dateCol, weekCol);
      寫入記憶(storageKey(cells), { dateCol: dateCol, weekCol: weekCol });
      $('manualPanel').classList.add('hidden');
      重新整理預覽();
    }

    function 取得表格() {
      return state.format === 'xlsx' ? state.parsed.table : state.parsed.grids[0];
    }

    function 重新整理預覽() {
      const grid = 取得表格();
      state.plan = WritePlan.buildWritePlan({
        monthModel: state.model, table: grid, headerRow: state.cols.headerRow,
        dateCol: state.cols.dateCol, weekCol: state.cols.weekCol, checkCols: state.cols.checkCols
      });

      state.yearMonthTargets = state.format === 'xlsx'
        ? ExcelAdapter.findYearMonthTargets(state.parsed.sharedStrings)
        : WordAdapter.findParagraphYearMonth(state.parsed.documentXml);

      渲染訊息();
      渲染年月清單();
      渲染預覽表(grid);
      $('outputName').textContent = '輸出為：' +
        outputFilename(state.file.name, state.year, state.month);
      $('previewPanel').classList.remove('hidden');
    }

    function 渲染訊息() {
      const parts = [];
      if (state.plan.missingDays.length > 0) {
        parts.push('<div class="warn"><strong>表格列數不足。</strong>這幾天寫不進去：' +
          state.plan.missingDays.join('、') + ' 日。仍可繼續，但這些日子不會出現在檔案裡。</div>');
      }
      if (state.plan.contentLoss.length > 0) {
        const items = state.plan.contentLoss.map(function (c) {
          return '第 ' + (c.row + 1) + ' 列：' + esc(c.text);
        }).join('<br>');
        parts.push('<div class="warn"><strong>下列內容將被清空。</strong><br>' + items + '</div>');
      }
      if (state.yearMonthTargets.length === 0) {
        parts.push('<div class="muted">未在文件中偵測到年月文字，將只填寫表格。</div>');
      }
      $('messages').innerHTML = parts.join('');
    }

    function 渲染年月清單() {
      const box = $('yearMonthList');
      if (state.yearMonthTargets.length === 0) { box.innerHTML = ''; return; }

      const rows = [];
      state.yearMonthTargets.forEach(function (t) {
        const 索引 = t.index !== undefined ? t.index : t.paraIndex;
        t.matches.forEach(function (m, i) {
          const key = 索引 + '-' + i;
          const 新值 = YearMonthText.rewriteMatch(m, state.year, state.month);
          rows.push('<label style="display:block"><input type="checkbox" class="ym" value="' +
            esc(key) + '" checked> ' + esc(m.raw) + ' → <strong>' + esc(新值) + '</strong></label>');
        });
      });
      box.innerHTML = '<div class="warn"><strong>文件中的年月將這樣改：</strong>' +
        rows.join('') + '</div>';
    }

    function 渲染預覽表(grid) {
      const 表頭 = grid[state.cols.headerRow] || [];
      const html = ['<table class="preview"><thead><tr><th>列</th>'];
      表頭.forEach(function (h) { html.push('<th>' + esc(h) + '</th>'); });
      html.push('</tr></thead><tbody>');

      const byCell = {};
      state.plan.writes.forEach(function (w) { byCell[w.row + ',' + w.col] = w.value; });

      state.plan.rowSummaries.forEach(function (s) {
        const cls = s.kind === 'holiday' ? 'is-holiday' : (s.kind === 'makeup' ? 'is-makeup' : '');
        html.push('<tr class="' + cls + '"><td>' + (s.row + 1) + '</td>');
        表頭.forEach(function (_, c) {
          const v = byCell[s.row + ',' + c];
          html.push('<td>' + esc(v) + '</td>');
        });
        html.push('</tr>');
      });
      html.push('</tbody></table>');
      $('preview').innerHTML = html.join('');
    }

    function 收集選取的年月鍵值() {
      const boxes = document.querySelectorAll('input.ym');
      const keys = [];
      for (let i = 0; i < boxes.length; i++) {
        if (boxes[i].checked) keys.push(boxes[i].value);
      }
      return keys;
    }

    function 下載() {
      const keys = 收集選取的年月鍵值();
      const 產生 = state.format === 'xlsx' ? 產生Excel() : 產生Word();
      產生.then(function (blob) {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = outputFilename(state.file.name, state.year, state.month);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
      }).catch(function (err) {
        顯示錯誤('產生檔案失敗：' + err.message);
      });

      function 產生Excel() {
        const applied = ExcelAdapter.applyWritesToSheet(
          state.parsed.sheetXml, state.plan.writes, state.parsed.cells);
        const ss = ExcelAdapter.rewriteSharedStrings(
          state.parsed.sharedStringsXml, state.parsed.sharedStrings,
          state.yearMonthTargets, state.year, state.month, keys);
        return ExcelAdapter.saveWorkbook(
          state.parsed.zip, state.parsed.sheetPath, applied.xml, ss, JSZip, 'blob');
      }

      function 產生Word() {
        let xml = WordAdapter.applyWritesToDocument(
          state.parsed.documentXml, state.parsed.tables[0], state.plan.writes);
        const targets = WordAdapter.findParagraphYearMonth(xml);
        xml = WordAdapter.rewriteParagraphYearMonth(xml, targets, state.year, state.month, keys);
        return WordAdapter.saveDocument(state.parsed.zip, xml, JSZip, 'blob');
      }
    }

    function 綁定拖放() {
      const drop = $('drop');
      const input = $('file');
      drop.addEventListener('click', function () { input.click(); });
      input.addEventListener('change', function () {
        if (input.files[0]) 讀取檔案(input.files[0]);
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
        if (ev.dataTransfer.files[0]) 讀取檔案(ev.dataTransfer.files[0]);
      });
    }

    初始化選單();
    綁定拖放();
    $('download').addEventListener('click', 下載);
    $('manualConfirm').addEventListener('click', 套用手動指定);
    重新整理月份();
  }

  return {
    outputFilename: outputFilename,
    detectFormat: detectFormat,
    storageKey: storageKey,
    columnsFrom: columnsFrom,
    start: start
  };
});
```

- [ ] **Step 4: 執行測試確認通過**

Run: `node run-tests.js`
Expected: PASS，新增的 14 項全數通過

- [ ] **Step 5: 在瀏覽器手動驗證整個流程**

用瀏覽器開啟 `src/index.html`（雙擊即可），依序確認：

1. 年月預設為當下的年月，不需手動選
2. 月曆顯示出來，週六日為紅色
3. 若能連網，資料來源顯示「已取得 20XX 年最新假日資料」
4. 把 `samples/機房月檢查表-範例.xlsx` 拖進去
5. 預覽出現，假日列標紅、檢查欄顯示 `-`
6. 警示區出現「空調異音已通報維護廠商」將被清空
7. 年月清單顯示 `115年8月 → 115年X月`
8. 按下下載，取得檔案後用瀏覽器重新拖入該檔，確認內容已更新且可正常讀取
9. 換 `.docx` 範例再走一次流程
10. 隨便拖一個 `.pdf` 進去，確認出現明確的錯誤訊息
11. **驗證手動指定**：用 Excel 或線上工具把範例檔的「日期」表頭改成「作業日」、
    「星期」改成「Wk」，讓自動偵測失敗。拖進工具後應出現「指定欄位」區塊，
    列出表格前幾列。指定表頭列與兩個欄位後按確定，預覽應正常出現。
12. **驗證記憶功能**：重新整理頁面，再拖入同一個檔案，
    這次應直接進入預覽，不再要求指定欄位

- [ ] **Step 6: 提交**

```bash
git add src/ui-main.js src/index.html tests/ui-main.test.js run-tests.js tests.html
git commit -m "新增介面主流程：讀檔、偵測、預覽、下載

舊格式明確擋下並提示另存新檔，
預覽區列出將被清空的文字內容與年月改寫對照。
欄位自動偵測失敗時改由使用者指定，並以表頭文字為鍵記憶於 localStorage。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 15: 合成單一發布檔與說明文件

把所有模組與 JSZip 內嵌成一個 `.html`，並寫給使用者看的說明。

**Files:**
- Create: `build.js`
- Create: `dist/檢查表星期填寫工具.html`（由 build 產生）
- Create: `README.md`

**Interfaces:**
- Consumes: `src/index.html` 及其載入的所有檔案
- Produces: 單一自足的 HTML 發布檔

**合成方式**：讀取 `src/index.html`，把每個 `<script src>` 換成該檔內容的行內版本，把 `<link rel=stylesheet>` 換成行內 `<style>`。**不做壓縮**——發布檔要能被人打開檢視，這在受管控的公司環境中是可稽核性的一部分。

- [ ] **Step 1: 寫 `build.js`**

```javascript
// 把 src/index.html 及其相依合成單一自足的 HTML 發布檔
// 用法：node build.js
const fs = require('fs');
const path = require('path');

const root = __dirname;
const srcDir = path.join(root, 'src');
const outDir = path.join(root, 'dist');
const outFile = path.join(outDir, '檢查表星期填寫工具.html');

function 讀取相對檔案(relative) {
  const p = path.resolve(srcDir, relative);
  if (!fs.existsSync(p)) throw new Error('找不到檔案：' + p);
  return fs.readFileSync(p, 'utf8');
}

let html = fs.readFileSync(path.join(srcDir, 'index.html'), 'utf8');

// 樣式表改為行內
html = html.replace(/<link\s+rel="stylesheet"\s+href="([^"]+)"\s*>/g, function (_, href) {
  console.log('內嵌樣式 ' + href);
  return '<style>\n' + 讀取相對檔案(href) + '\n</style>';
});

// 腳本改為行內
html = html.replace(/<script\s+src="([^"]+)"><\/script>/g, function (_, src) {
  console.log('內嵌腳本 ' + src);
  const code = 讀取相對檔案(src);
  // 避免程式碼中出現的結束標籤提前關閉 script 區塊
  return '<script>\n' + code.replace(/<\/script>/gi, '<\\/script>') + '\n</script>';
});

const 標記 = '<!-- 由 build.js 合成，請勿直接編輯。原始碼在 src/ 目錄。 -->\n';
html = html.replace('</head>', 標記 + '</head>');

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, html, 'utf8');

const kb = Math.round(fs.statSync(outFile).size / 1024);
console.log('');
console.log('已產生 ' + outFile + '（' + kb + ' KB）');
if (html.indexOf('<script src=') !== -1 || html.indexOf('<link rel="stylesheet"') !== -1) {
  throw new Error('仍有未內嵌的外部檔案，發布檔不是自足的');
}
console.log('確認：檔案中已無任何外部相依');
```

- [ ] **Step 2: 執行合成**

Run: `node build.js`
Expected: 逐行印出內嵌的檔案，最後顯示產生的檔案大小（約 250 至 400 KB，主要來自 JSZip 與內建假日資料），並印出「確認：檔案中已無任何外部相依」

- [ ] **Step 3: 驗證發布檔確實自足**

Run:
```bash
node -e "
const s=require('fs').readFileSync('dist/檢查表星期填寫工具.html','utf8');
const 外部 = s.match(/(src|href)=\"(?!#)[^\"]*\"/g) || [];
console.log('外部參照：', 外部.length === 0 ? '無' : 外部.join(', '));
console.log('含 JSZip：', s.indexOf('JSZip') !== -1);
console.log('含內建假日：', s.indexOf('HolidaysBuiltin') !== -1);
console.log('含 20260925：', s.indexOf('20260925') !== -1);
"
```
Expected: 外部參照為「無」，其餘三項皆為 `true`

- [ ] **Step 4: 在瀏覽器對發布檔做端對端驗證**

**把 `dist/檢查表星期填寫工具.html` 複製到另一個資料夾**（例如桌面）再開啟——這一步是為了證明它不依賴專案目錄裡的任何檔案。

依序確認：

1. 頁面正常顯示，年月預設為當下年月
2. 月曆正確、假日標紅
3. 拖入 `samples/機房月檢查表-範例.xlsx`，預覽正確
4. 下載後的檔案再拖回工具，內容正確且可讀
5. `.docx` 範例同樣走一次
6. **中斷網路後重新整理頁面**，確認資料來源顯示為「使用內建的 20XX 年假日資料（無法連線）」，且整個流程仍可完成——這是離線可用性的關鍵驗證

- [ ] **Step 5: 寫 `README.md`**

````markdown
# 機房月檢查表　星期填寫工具

依所選年月自動填寫機房月檢查表的日期與星期，假日整列填入 `-`，並同步文件標題中的年月。

支援 Excel（`.xlsx`）與 Word（`.docx`）。

## 怎麼用

1. 下載 `dist/檢查表星期填寫工具.html`
2. 雙擊用瀏覽器打開（不需要安裝任何東西）
3. 確認年月（預設就是當下的年月）
4. 把檢查表拖進去
5. 看過預覽沒問題，按下下載

改好的檔案會是**新檔名**（例如 `機房檢查表_115年09月.xlsx`），原檔不會被動到。

## 它會做什麼

| 情況 | 日期欄 | 星期欄 | 檢查項目欄 |
|---|---|---|---|
| 平日 | 填入日數 | 填入星期 | 清空 |
| 週六日、國定假日 | 填入日數 | 填入星期 | 全部填 `-` |
| 補班日（週六要上班） | 填入日數 | 填入星期 | 視同平日，清空 |
| 超過當月天數的列 | 清空 | 清空 | 清空 |

假日判斷含國定假日、彈性放假與補班日。有網路時自動取得最新資料；沒網路時使用內建資料（涵蓋 2024 至 2027 年）；連內建資料都沒有的年份，只依週六日判斷——畫面上一定會標明目前用的是哪一種。

颱風假之類的臨時狀況，在月曆上點一下那天就能切換。

## 注意事項

- 只支援 `.xlsx` 與 `.docx`。舊的 `.xls` / `.doc` 請先用 Office 另存新檔。
- **平日的檢查欄會被清空。** 如果拖進去的是已經填好的檔案，填過的紀錄會不見。預覽會先列出即將被清空的文字內容，原檔則一律不動。
- 表格必須有「日期」和「星期」兩個欄位標題，工具靠它們定位。

## 開發

```bash
npm install              # 安裝開發用套件
npm test                 # 跑測試（也可以直接雙擊 tests.html）
npm run samples          # 產生仿真範例檔
npm run holidays         # 更新內建假日資料
npm run build            # 合成 dist/ 的單一檔案
```

原始碼在 `src/`，發布檔由 `build.js` 合成，**不要直接編輯 `dist/` 裡的檔案**。

政府每年底公布次年度行事曆，屆時執行 `npm run holidays` 與 `npm run build` 更新內建資料。

設計文件與實作計畫在 `docs/superpowers/`。
````

- [ ] **Step 6: 執行完整測試確認全綠**

Run: `node run-tests.js`
Expected: 全部通過，失敗 0 項

- [ ] **Step 7: 提交**

`dist/` 的成品必須進版控——使用者端沒有 Node，只能直接下載成品使用。

```bash
git add build.js dist/ README.md
git commit -m "新增發布合成腳本與使用說明

把所有模組與 JSZip 內嵌成單一自足的 HTML，不做壓縮以保留可讀性。
發布檔納入版控，供無開發環境的使用者直接下載使用。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## 完成標準

全部任務完成後，下列每一項都必須成立：

- `node run-tests.js` 全數通過，失敗 0 項
- 雙擊 `tests.html` 在瀏覽器中同樣全數通過
- `dist/檢查表星期填寫工具.html` 複製到任意資料夾後仍能正常運作
- 中斷網路後該檔仍可完成整個流程，且畫面標明使用內建資料
- Excel 與 Word 範例檔都能正確填寫並可重新讀回
- 改寫後的 Excel `xl/styles.xml` 與原檔逐字相同（框線未受影響）
- 改寫後的 Word 表格框線設定仍在，標題粗體仍在
- 原始檔案在任何情況下都未被修改
- 檔案內容經跳脫後才進入畫面：把某格內容改成 `<img src=x onerror=alert(1)>` 後拖入工具，
  預覽應原樣顯示該段文字，不得跳出對話框
