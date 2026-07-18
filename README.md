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
├── index.html   # 主表單：付款申請單（填單 + A4 預覽 + 列印 + 歷史資料清單）
├── proof.html   # 輔助表單：支出證明單（無單據支出用；存檔+列印+歷史資料）
├── food.html    # 輔助表單：伙食費支出證明單（存檔+列印+歷史資料）
├── Code.gs      # GAS 後端：JSON API（科目/流水號/存檔/查詢/輔助表單），需手動貼進 Apps Script
├── .nojekyll    # 停用 GitHub Pages 的 Jekyll 建置
└── README.md
```
三個頁面左上角有互相切換的連結。

## 核心設計（定案紀錄）

- **存檔即會計格式**：資料直接寫入每月分頁 `會計支出115MM`（自動建立），
  **B～J 欄與會計原始 xlsx 完全相同**，沒有中間資料表、沒有「產月報」步驟。
- **一張單可多科目，一筆明細一列**（2026-07 改版：**不再自動合併同科目**）：
  每筆明細各自一列寫入，科目各自對照，金額為單筆數字（不再寫 `=a+b` 公式）。
- **NO／流水號：暫停自動編號**（手工、電子並行期間會錯亂）：
  NO 欄留空白，前端 placeholder「尚未開放自動編號」；使用者手動輸入的話照存。
  自動編號程式（`maxNo_`/`peekNextNo`）保留在 Code.gs，將來要恢復把
  `doGet` 的 nextno 改回 `peekNextNo()`、前端加回 `suggestNo()` 即可。
- **列印自動存檔**：主按鈕「存檔並列印」未存檔或有改動就先存再印；空白單直接印。
  「只存檔」給不印紙本的情況；「歷史資料」查詢/重印。
- **重印不重複存檔**：從「歷史資料」載入後 `dirty=false`，直接列印不再寫入；
  修改內容後列印則以新 NO（或新的一列）另存一筆。
- **歷史資料三頁都有**：三頁都靠存檔時多存的「**原始資料**」欄（完整 JSON）無損還原表單
  （付款申請單存在每張單第一列的 L 欄）。沒有原始資料的舊列退回粗略還原（依 NO 分組、
  會計名稱當科目），科目名稱可能與當初不同。
- **列印無頁首頁尾**：`@page{margin:0}` 讓瀏覽器沒空間印標題/網址/頁碼，
  紙張留白由 `.sheet` 的 padding 提供（三頁同做法）。
- **導覽**：三頁左上角膠囊分頁籤互相切換（當前頁深色）；主頁按鈕區另有
  兩顆連結按鈕直達證明單。
- **簽名**：前端只填「申請人」（必填），Sheet 的「領款人」欄記申請人姓名；
  A4 列印的領款人簽名／申請人一律留空給實體簽名。
- **科目**：「科目」分頁**三欄**——A 科目名稱、B 會計科目對照（如 `清掃薪資 → 6211薪資費用`）、
  C 說明。前端下拉顯示「**科目（說明）**」幫助選擇，但存檔與 PDF 只用科目名。
  存檔時映射成「代號+名稱」；**未對照到的列會整列標黃**，提醒補對照。
- **金額位數格**：最高「拾萬」（上限 999,999），超過前端跳紅字提醒拆單。
- 會計原始 xlsx（`會計格式現支115年度*.xlsx`）**不進 git**（.gitignore 已排除），
  範本檔在維護者本機 `Downloads\liming-voucher\`。

## Google Sheet 結構

- **「科目」**（設定頁）：A 欄科目名稱（前端下拉）、B 欄會計科目（代號+名稱）、
  C 欄說明（下拉顯示「科目（說明）」用，PDF 不印）。
  試算表選單「黎明教會 → 初始化科目分頁」可建 11 個預設對照（已有資料不覆寫）。
- **「會計支出115MM」**（每月資料頁，存檔自動建立）：
  - 標題在第 2 列、B 欄起（比照會計 xlsx 版面）：
    `NO | 月 | 日期 | 科目代號 | 會計名稱 | 摘要 | 合計 | 領款人 | 備註 | 建立時間 | 原始資料`
  - K（建立時間）、L（原始資料 JSON，歷史還原用、**勿刪改**）是系統額外欄，
    交給會計前可隱藏或刪除（刪 L 該單就無法從歷史資料還原）
  - NO 欄目前留空（自動編號暫停），手寫編號者可自行填入
  - 日期為民國文字（`115/07/05`）；會計名稱為 `=MID(E?,5,LEN(E?))` 公式
  - 最下方合計列（`=SUM(H3:H?)`＋I 欄「合計」）由程式在每次存檔後自動維護
  - 備註欄（J）由出納/會計直接在 Sheet 上註記（如「06/20已給會計」）
  - **不要手動改「合計」那一列**；改資料列沒問題，程式只認 I 欄=「合計」的列
- **「支出證明單」「伙食費支出證明單」**（輔助表單流水帳，存檔自動建立）：
  一張表單一列，明細序列化成文字（如 `白米×2包=1500；蔬菜×一批=2935`），
  無特定格式要求，純留存紀錄；正式文件以列印的紙本（含簽名）為準。
  最後一欄「**原始資料**」存完整 JSON——歷史資料還原表單就是靠它，**請勿刪除或改動**。
  標題列每次存檔會重寫（冪等），未來加欄位舊分頁會自動補標題。

## 後端 API（Code.gs）

Base URL：GAS Web App 的 `/exec` 網址（已填在 `index.html` 的 `GAS_API_URL`）。
前後端用 `text/plain` POST 溝通以避開 CORS preflight。

| 方法 | 參數 | 說明 |
|---|---|---|
| GET `?action=subjects` | — | 科目陣列 `[{name,desc}]`（下拉顯示 名稱（說明）） |
| GET `?action=nextno` | — | **已暫停**，固定回 `{"no":""}` |
| GET `?action=list` | `limit`(預設20)、`month`(選填,如11506) | 該月資料列陣列（HEADER 為 key） |
| GET `?action=prooflist` | `limit`(預設50) | 支出證明單流水帳（含原始資料 JSON，給歷史還原） |
| GET `?action=foodlist` | `limit`(預設50) | 伙食費支出證明單流水帳（含原始資料 JSON） |
| POST `action:"save"` | body 見下 | 付款申請單存檔 |
| POST `action:"proof"` | rocDate/subject/reason/why/payee/items/total | 支出證明單存檔（一單一列流水帳） |
| POST `action:"food"` | rocDate/reason/from/to/payee/people/tables/meals/buys/receipts/total | 伙食費支出證明單存檔（一單一列流水帳） |

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
2. 依下方「**GAS 部署教學**」的情境二更新部署（貼程式 → 儲存 → 新版本 → 部署）

### 本機開發
- `GAS_API_URL` 留空字串＝示範模式（預設科目、存檔只印 Console），方便改 UI
- repo 附 `.claude/launch.json`，`python -m http.server 8123` 也可

### 快速驗證後端版本
```bash
curl -sL "<EXEC_URL>?action=nextno"     # 應回 {"no":"115MM-NN"}
curl -sL "<EXEC_URL>?action=subjects"   # 應回科目陣列
```

## GAS 部署教學（手把手）

先搞懂三個名詞，之後的步驟就不會迷路：

- **Apps Script 專案**：放 `Code.gs` 程式碼的地方。本專案的叫「Voucher出納」，
  **綁定**在試算表上（從試算表的「擴充功能 → Apps Script」進入）。
- **部署（Web App）**：把程式碼發布成一個對外網址。**存檔（Ctrl+S）不等於部署**——
  存檔只是更新程式碼草稿，線上跑的永遠是「最後一次部署的版本」。
- **`/exec` 網址**：部署後得到的 API 網址，填在 `index.html` 最上面的 `GAS_API_URL`。

### 情境一：第一次從零建置（例如換一個全新的試算表）

1. 建一個 Google Sheet（用哪個帳號建，那個帳號就是之後唯一能管部署的人，請慎選）
2. 在試算表裡點「**擴充功能 → Apps Script**」
   ⚠️ 一定要從試算表裡進去，程式才會「綁定」這張表；
   直接去 script.google.com 建的「獨立專案」抓不到試算表，會報
   `TypeError: Cannot read properties of null (reading 'getSheetByName')`
3. 把 repo 裡 `Code.gs` **全文**貼上（蓋掉原本的空函式）→ `Ctrl+S` 儲存
4. 右上角「**部署 → 新增部署作業**」→ 左上齒輪選「**網頁應用程式**」
5. 設定兩個關鍵欄位：
   - 執行身分：**我**
   - 誰可以存取：**任何人**（前端網頁沒登入 Google 也要能呼叫，一定要選這個）
6. 按「部署」→ 第一次會跳授權流程：
   選帳號 → 出現「**Google 尚未驗證這個應用程式**」警告（自己寫的程式都會這樣，正常）
   → 點「**進階**」→「**前往 Voucher出納（不安全）**」→「**允許**」
7. 完成後複製「網頁應用程式」的網址（結尾是 `/exec`）
8. 貼進 `index.html` 的 `GAS_API_URL` → commit → push（等 Pages 部署完成）
9. 回試算表按 F5 重新整理 → 工具列會多出「**黎明教會**」選單 → 點「**初始化科目分頁**」
10. 驗證：瀏覽器開 `<exec網址>?action=nextno`，看到 `{"no":"115MM-01"}` 就通了

### 情境二：更新程式碼（日常最常用）

1. 打開 Apps Script（試算表 → 擴充功能 → Apps Script）
2. 把新版 `Code.gs` 全文貼上蓋掉舊的 → `Ctrl+S` 儲存
3. **「部署 → 管理部署作業」→ 點鉛筆圖示 → 「版本」下拉選「新版本」→ 按「部署」**
   （這步才是真正上線；網址不變，前端不用改）
4. 等 1～2 分鐘再測試（Google 端有快取，剛部署完可能新舊版短暫並存）

⚠️ 第 3 步**只有部署建立者**（目前是 `lmf@tjcedu.org`）能按，
其他帳號即使有試算表編輯權，點「管理部署作業」也會跳「發生錯誤」。
如果將來要換人維護，最乾淨的做法是由新維護者用「新增部署作業」建自己的部署，
拿到新的 `/exec` 網址後更新 `GAS_API_URL`。

### 常見錯誤對照表

| 症狀 | 原因 | 解法 |
|---|---|---|
| API 回 `TypeError: Cannot read ... null` | 專案是「獨立專案」，沒綁定試算表 | 從試算表的「擴充功能 → Apps Script」重建（情境一） |
| 「管理部署作業」跳紅色「發生錯誤」 | 不是部署建立者的帳號 | 換 `lmf@tjcedu.org` 操作，或建自己的新部署 |
| 改了 Code.gs 但行為沒變 | 只存檔沒部署（存檔≠部署） | 做情境二的第 3 步「新版本 → 部署」 |
| 存進 Sheet 的資料欄位錯位／科目空白 | 線上還是舊版程式，或剛部署完新舊並存 | 確認已部署新版本；錯位的列手動刪掉重存 |
| `/exec` 開起來是 Google 錯誤頁 | 授權沒完成，或「誰可以存取」不是「任何人」 | 重新部署並跑完授權流程，存取權選「任何人」 |
| 試算表沒有「黎明教會」選單 | 綁定後還沒重新整理，或 onOpen 沒觸發 | 試算表按 F5；還是沒有就檢查程式是否貼齊全 |

## 給接手者的程式重點

- **index.html**（全部邏輯在 `<script>` 內）：
  - `state`＝資料模型（no/date/applicant/items[]，每列明細含 subject/memo/amount）
  - `renderItems()` 畫左側明細編輯、`renderSheet()` 畫右側 A4 預覽
  - `splitDigits()` 金額拆位數格、`COLS` 定義位數欄（要加「百萬」欄改這裡＋上限）
  - `dirty` 旗標＋`printSheet()`＝列印前自動存檔；`saveVoucher()` 回傳 true/false
  - `openList()/loadVoucher()`＝歷史資料清單與載入重印（`LIST_CACHE` 依 NO 分組）
  - 金額輸入「輸入中不重畫、失焦才補千分位」——別改回每鍵 render，會掉焦點
  - 明細欄位樣式選擇器是 `.item .grid input/select`（不在 `.field` 底下）
- **proof.html / food.html**（各自獨立單檔，結構仿 index）：
  - `openList()/loadRecord()`＝歷史資料：抓 `prooflist`/`foodlist`，
    解析每列的「原始資料」JSON 還原 state（`rocToIso()` 把民國轉回 date input 格式）
  - proof：數量×單價自動算總價（可手改）；`cnDigits()` 合計中文位數（萬仟佰拾元）
  - food：`twoCol()` 把清單拆左右兩欄呈現；用餐情形合計列左右各自加總人數與金額
- **Code.gs**：
  - `saveVoucher()`：科目分組 → 寫入月分頁 → 補 MID 公式與合計列；全程 LockService
  - `getMonthSheet_()` 自動建月分頁含格式；`findSumRow_()` 以 I 欄=「合計」定位
  - `maxNo_()` 取月內最大 NO；前端帶來的 no 若 ≤ maxNo 自動遞補
  - `saveProof()/saveFood()` → `appendSimple_()`（標題列冪等重寫）；
    `listSimple_()` 回傳輔助表單流水帳（含原始資料欄）
  - 所有民國日期／期間欄寫入都帶 `'` 前綴強制文字，否則 Sheet 會誤判成西元 115 年

## 待辦 / 可能的後續

- [ ] 試算表更名（目前叫「付款申請單(測試中)」，正式啟用後改名，程式不受影響）
- [ ] （可選）跨月查詢頁：`list` API 已支援 `month` 參數，前端加月份選單即可
- [ ] （可選）存檔同時自動產生 PDF 存 Drive／寄出（需 GAS 端做 Sheet→PDF）
- [ ] （可選）科目管理 UI（目前直接編輯「科目」分頁，夠用）
- [ ] （可選）刪單/改單功能（目前直接在 Sheet 刪列；注意別動合計列）
