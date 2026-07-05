# liming-voucher — 黎明教會付款申請單（代傳票）

真耶穌教會 黎明教會的付款申請單系統：網頁填單 → 存進 Google Sheet（即會計 Excel 格式）→ 列印 A4 代傳票 → 月底會計直接下載。

## 線上資源

| 項目 | 位置 |
|---|---|
| 前端（GitHub Pages） | https://tjcliming.github.io/liming-voucher/ |
| GitHub repo | https://github.com/TJCLiMing/liming-voucher |
| Google Sheet | 「付款申請單(測試中)」，擁有者 `lmf@tjcedu.org` |
| GAS 專案 | 「Voucher出納」，綁定在上述試算表（擴充功能 → Apps Script） |
| Web App 部署 | 由 `lmf@tjcedu.org` 建立——**只有這個帳號能更新部署**（見維運 SOP） |

## 整體流程

```
申請人開網頁填單（每列明細：科目＋摘要＋金額，申請人必填）
   → 按「存檔並列印」：自動存進當月分頁「會計支出115MM」、取得單號、跳列印
   → A4 代傳票（金額拆位數格「拾萬～元」；領款人/申請人留空給實體簽名）
   → 月底會計直接下載 Google Sheet，即為最終 Excel 格式，零轉換
```

## 檔案結構

```
liming-voucher/
├── index.html   # 前端：填單 UI + 即時 A4 預覽 + 列印 + 歷史資料清單（單檔 vanilla JS）
├── Code.gs      # GAS 後端：JSON API（科目/流水號/存檔/查詢），需手動貼進 Apps Script
├── .nojekyll    # 停用 GitHub Pages 的 Jekyll 建置
└── README.md
```

## 核心設計（定案紀錄）

- **存檔即會計格式**：資料直接寫入每月分頁 `會計支出115MM`（自動建立），
  **B～J 欄與會計原始 xlsx 完全相同**，沒有中間資料表、沒有「產月報」步驟。
- **一張單可多科目**（仿會計 xlsx 11506 的 NO16）：
  - 同科目多筆明細 → 併成一列：摘要「、」串接、合計寫 `=a+b+c` 公式（比照會計手工習慣）
  - 不同科目 → 各一列，**共用同一個 NO**
- **NO**：月內流水整數（1、2、3…）；前端顯示/紙本印 `11507-16`（民國年月-NO）。
  前端建議號若被同時段其他人用走，存檔時自動遞補下一號。
- **列印自動存檔**：主按鈕「存檔並列印」未存檔或有改動就先存再印；空白單直接印。
  「只存檔」給不印紙本的情況；「歷史資料」查詢/重印。
- **重印不重複存檔**：從「歷史資料」載入後 `dirty=false`，直接列印不再寫入；
  修改內容後列印則以新 NO 另存一筆。
- **簽名**：前端只填「申請人」（必填），Sheet 的「領款人」欄記申請人姓名；
  A4 列印的領款人簽名／申請人一律留空給實體簽名。
- **科目**：前端下拉用自訂名稱（伙食費、雜費…），「科目」分頁 B 欄放會計科目對照
  （如 `清掃薪資 → 6211薪資費用`）。存檔時映射成「代號+名稱」；
  **未對照到的列會整列標黃**，提醒到「科目」分頁補對照。
- **金額位數格**：最高「拾萬」（上限 999,999），超過前端跳紅字提醒拆單。
- 會計原始 xlsx（`會計格式現支115年度*.xlsx`）**不進 git**（.gitignore 已排除），
  範本檔在維護者本機 `Downloads\liming-voucher\`。

## Google Sheet 結構

- **「科目」**（設定頁）：A 欄科目名稱（前端下拉）、B 欄會計科目（代號+名稱）。
  試算表選單「黎明教會 → 初始化科目分頁」可建 11 個預設對照（已有資料不覆寫）。
- **「會計支出115MM」**（每月資料頁，存檔自動建立）：
  - 標題在第 2 列、B 欄起（比照會計 xlsx 版面）：
    `NO | 月 | 日期 | 科目代號 | 會計名稱 | 摘要 | 合計 | 領款人 | 備註 | 建立時間`
  - K（建立時間）是系統額外欄，交給會計前可隱藏或刪除
  - 日期為民國文字（`115/07/05`）；會計名稱為 `=MID(E?,5,LEN(E?))` 公式
  - 最下方合計列（`=SUM(H3:H?)`＋I 欄「合計」）由程式在每次存檔後自動維護
  - 備註欄（J）由出納/會計直接在 Sheet 上註記（如「06/20已給會計」）
  - **不要手動改「合計」那一列**；改資料列沒問題，程式只認 I 欄=「合計」的列

## 後端 API（Code.gs）

Base URL：GAS Web App 的 `/exec` 網址（已填在 `index.html` 的 `GAS_API_URL`）。
前後端用 `text/plain` POST 溝通以避開 CORS preflight。

| 方法 | 參數 | 說明 |
|---|---|---|
| GET `?action=subjects` | — | 科目名稱陣列（給下拉） |
| GET `?action=nextno` | — | 建議單號 `{"no":"11507-03"}` |
| GET `?action=list` | `limit`(預設20)、`month`(選填,如11506) | 該月資料列陣列（HEADER 為 key） |
| POST | body 見下 | 存檔 |

POST body（`action:"save"`）：
```json
{ "action":"save", "no":"11507-03 或空字串", "date":"2026-07-05",
  "rocDate":"115-07-05", "applicant":"姓名", "total":6633,
  "items":[ {"subject":"伙食費","memo":"...","amount":4618}, ... ] }
```
回應：`{"ok":true,"no":"11507-03","rows":2,"unmapped":0}`

## 維運 SOP

### 改前端
1. 改 `index.html` → commit → push，GitHub Pages 自動部署（約 1 分鐘）
2. **已知問題**：Pages 的 deploy 步驟偶爾無故失敗（build 成功、deploy 失敗），
   頁面會停在舊版。確認方式：repo → Actions 看最新 run；失敗就隨便推個 commit 重觸發
3. 瀏覽器記得 `Ctrl+Shift+R` 避開快取

### 改後端
1. 改 `Code.gs`（repo 裡的是唯一正本）→ commit → push
2. 全文貼進 Apps Script 專案「Voucher出納」→ 儲存
3. **用 `lmf@tjcedu.org` 帳號**：部署 → 管理部署作業 → 鉛筆 → 版本「新版本」→ 部署
   （網址不變；其他帳號按會跳權限錯誤，Apps Script 規定只有部署建立者能更新）
4. **已知問題**：重新部署後幾十秒內新舊版可能並存（Google 端快取），
   期間存檔可能寫出欄位錯位的列，發現就手動刪掉該列重存

### 本機開發
- `GAS_API_URL` 留空字串＝示範模式（預設科目、存檔只印 Console），方便改 UI
- repo 附 `.claude/launch.json`，`python -m http.server 8123` 也可

### 快速驗證後端版本
```bash
curl -sL "<EXEC_URL>?action=nextno"     # 應回 {"no":"115MM-NN"}
curl -sL "<EXEC_URL>?action=subjects"   # 應回科目陣列
```

## 給接手者的程式重點

- **index.html**（全部邏輯在 `<script>` 內）：
  - `state`＝資料模型（no/date/applicant/items[]，每列明細含 subject/memo/amount）
  - `renderItems()` 畫左側明細編輯、`renderSheet()` 畫右側 A4 預覽
  - `splitDigits()` 金額拆位數格、`COLS` 定義位數欄（要加「百萬」欄改這裡＋上限）
  - `dirty` 旗標＋`printSheet()`＝列印前自動存檔；`saveVoucher()` 回傳 true/false
  - `openList()/loadVoucher()`＝歷史資料清單與載入重印（`LIST_CACHE` 依 NO 分組）
  - 金額輸入「輸入中不重畫、失焦才補千分位」——別改回每鍵 render，會掉焦點
  - 明細欄位樣式選擇器是 `.item .grid input/select`（不在 `.field` 底下）
- **Code.gs**：
  - `saveVoucher()`：科目分組 → 寫入月分頁 → 補 MID 公式與合計列；全程 LockService
  - `getMonthSheet_()` 自動建月分頁含格式；`findSumRow_()` 以 I 欄=「合計」定位
  - `maxNo_()` 取月內最大 NO；前端帶來的 no 若 ≤ maxNo 自動遞補

## 待辦 / 可能的後續

- [ ] 試算表更名（目前叫「付款申請單(測試中)」，正式啟用後改名，程式不受影響）
- [ ] （可選）跨月查詢頁：`list` API 已支援 `month` 參數，前端加月份選單即可
- [ ] （可選）存檔同時自動產生 PDF 存 Drive／寄出（需 GAS 端做 Sheet→PDF）
- [ ] （可選）科目管理 UI（目前直接編輯「科目」分頁，夠用）
- [ ] （可選）刪單/改單功能（目前直接在 Sheet 刪列；注意別動合計列）
