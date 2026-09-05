# CLAUDE.md — 開發規範

本專案為「機房月檢查表星期自動填寫工具」，產出物是一個雙擊即開的單一自足 HTML 檔。

## 設計文件

- 需求規格：`docs/superpowers/specs/2026-08-30-week-automodify-design.md`
- 實作計畫（15 Task）：`docs/superpowers/plans/2026-08-30-week-automodify.md`
- 實作時以計畫為主、spec 為依據，有衝突以 spec 為準。

## 相依套件

- **唯一允許的第三方函式庫：JSZip**（3.10.1）。
- **明確禁用 SheetJS**——社群版寫出 `.xlsx` 不保留儲存格樣式，框線會遺失。
- 禁止引入任何其他套件。CDN 僅限假日資料抓取（jsDelivr），不載入 JS/CSS。

## 模組寫法

所有 `src/` 下的模組必須使用 UMD 模式，同時支援 Node `require` 與瀏覽器全域變數：

```javascript
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ModuleName = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  // ...
  return { /* 公開 API */ };
});
```

## Excel 寫入規則

- 寫入一律使用**行內字串** `t="inlineStr"` 或數字，**不得新增 `xl/sharedStrings.xml` 的項目**。
- 清空儲存格時**保留 `<c>` 元素及其 `s` 樣式屬性**，僅移除內部 `<v>` 或 `<is>`，確保框線與底色不消失。
- 改寫後 `styles.xml` 的內容必須與原檔完全相同（位元組一致）。

## Word 寫入規則

- 年月文字可能跨多個 `<w:r>` run，偵測時必須先合併整段再比對。
- 改寫時只動被觸及的 run，不拆不合 run 本身的格式屬性（`<w:rPr>`）。

## 寫入規則（業務邏輯）

| 情境 | 日期欄 | 星期欄 | 檢查項目欄 |
|---|---|---|---|
| 平日 | 寫入 1~N | 寫入對應星期 | **清空** |
| 假日（六日、國定假日） | 寫入日期 | 寫入對應星期 | **全部填 `-`** |
| 補班日（六日但要上班） | 寫入日期 | 寫入對應星期 | 視同平日，**不填 `-`** |
| 超出當月天數的列 | 清空 | 清空 | 清空 |

- 星期字串固定使用單字：`一 二 三 四 五 六 日`。
- 星期來源**優先採用假日資料的 `week` 欄位**，僅在第 3 層 fallback 才由程式推算。
- 「檢查項目欄」= 日期欄與星期欄以外的所有欄位，含備註欄。

## 假日資料

三層遞降：線上（jsDelivr）→ 內建（2024–2027）→ 純計算（僅六日）。
線上來源：`https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/<年>.json`

## 年月偵測

- 正規式：`([0-9]{2,4})[ ]*年[ ]*([0-9]{1,2})[ ]*月`
- 年份 ≤ 200 視為民國年，≥ 1911 視為西元年。
- 改寫時沿用原文的紀年制、補零、空白寫法，只換數字。

## 檔案處理

- **原檔永不覆蓋**，輸出新檔名格式：`原檔名_115年09月.xlsx`。
- 一次處理一個檔案。
- 只改日期欄、星期欄、檢查項目欄的值，以及使用者確認的年月文字，其餘一概不動。

## 預覽

- 預覽步驟不可省略。
- 被清空的儲存格若內容不是單一記號（V、X、v、x、○、✓、✗、-、空白），須醒目標示原文。

## 安全

- 來自檔案的文字放進 HTML 前必須跳脫（`XmlText.escape`），防止 XSS。

## 語言與風格

- 程式碼註解、UI 文字、commit 訊息一律繁體中文。
- commit 訊息結尾附上 `Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>`。

## 測試

- 日曆核心採 TDD。
- 測試同時在 Node（`node run-tests.js`）和瀏覽器（`tests.html`）執行。
- 測試檔使用與模組相同的 UMD 開頭模式。

## 發布

- `build.js` 合成單一 HTML 到 `dist/檢查表星期填寫工具.html`。
- 合成後的成品 commit 進 repo，使用者端不需要 Node。

## 禁止事項

- 不支援 `.xls` / `.doc` 舊格式。
- 不做批次處理。
- 不做假日列標色（只填 `-`，不動格式）。
- 不做線上部署或伺服器。
- 不加超出計畫範圍的功能。
