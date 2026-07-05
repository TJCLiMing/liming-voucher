/*************************************************************
 * 黎明教會 付款申請單（代傳票）後端 API
 * 部署：部署 → 新增部署作業 → 網頁應用程式
 *   執行身分：我
 *   誰可以存取：任何人
 * 把 /exec 網址貼到前端 index.html 的 GAS_API_URL
 *
 * 設計：存檔直接寫成會計 Excel 格式，每月一個分頁「會計支出115MM」
 *   B~J 欄與會計 xlsx 完全相同：
 *   NO | 月 | 日期 | 科目代號 | 會計名稱 | 摘要 | 合計 | 領款人 | 備註
 *   （領款人欄記「申請人」姓名；前端已無獨立領款人欄位）
 *   K 為系統額外欄：建立時間（會計可忽略或隱藏）
 *
 *   一張申請單可含多個科目：
 *   - 同科目多筆明細 → 併成一列，摘要用「、」串接，合計寫 =a+b 公式
 *   - 不同科目 → 各一列，共用同一個 NO（同會計 xlsx 的 NO16 做法）
 *   NO 為月內流水整數；前端顯示/紙本印「11507-16」
 *
 * 另需「科目」分頁：A 欄科目名稱（前端下拉）、B 欄會計科目（代號+名稱）
 *************************************************************/

const SHEET_SUBJECT = '科目';
const MONTH_PREFIX  = '會計支出';
// 月分頁欄位（從 B 欄起）：B=NO C=月 D=日期 E=科目代號 F=會計名稱 G=摘要 H=合計 I=領款人 J=備註 K=建立時間
const HEADER = ['NO','月','日期','科目代號','會計名稱','摘要','合計','領款人','備註','建立時間'];
const DATA_START_ROW = 3;   // 比照會計 xlsx：標題在第 2 列、B 欄起
const COL_B = 2, N_COLS = HEADER.length;

/* ---------- 試算表選單 ---------- */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('黎明教會')
    .addItem('初始化科目分頁','initSubjectSheet')
    .addToUi();
}

/* ---------- 路由 ---------- */
function doGet(e){
  const action = (e.parameter.action || '').toLowerCase();
  if(action === 'subjects') return json(getSubjects());
  if(action === 'nextno')   return json({ no: peekNextNo() });
  if(action === 'list')     return json(listRecent(Number(e.parameter.limit) || 20, e.parameter.month));
  return json({ ok:true, msg:'黎明教會 付款申請單 API' });
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.action === 'save')  return json(saveVoucher(body));
    if(body.action === 'proof') return json(saveProof(body));
    if(body.action === 'food')  return json(saveFood(body));
    return json({ ok:false, error:'unknown action' });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

/* ---------- 科目 ---------- */
function getSubjects(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUBJECT);
  if(!sh) return [];
  const last = sh.getLastRow();
  if(last < 2) return [];
  return sh.getRange(2,1,last-1,1).getValues()
           .map(r => String(r[0]).trim()).filter(Boolean);
}

/* 科目名稱 → 會計科目（代號+名稱，如 6331伙食費） */
function getSubjectMap_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUBJECT);
  const map = {};
  if(!sh) return map;
  const last = sh.getLastRow();
  if(last < 2) return map;
  sh.getRange(2,1,last-1,2).getValues().forEach(r=>{
    const name = String(r[0]).trim(), acct = String(r[1]).trim();
    if(name && acct) map[name] = acct;
  });
  return map;
}

/* 預設科目與會計科目對照，僅在「科目」分頁不存在或為空時建立 */
function initSubjectSheet(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SUBJECT);
  if(sh && sh.getLastRow() > 1){
    SpreadsheetApp.getUi().alert('「科目」分頁已有資料，未變更。');
    return;
  }
  if(!sh) sh = ss.insertSheet(SHEET_SUBJECT);
  const rows = [
    ['科目','會計科目'],
    ['伙食費','6331伙食費'],
    ['旅運費','6251旅運費'],
    ['清掃薪資','6211薪資費用'],
    ['文具費','6231文具用品'],
    ['郵電費','6261郵電費'],
    ['水電費','6291水電瓦斯費'],
    ['修繕費','6271修繕費'],
    ['誌謝金','6351接待費'],
    ['餐費','6331伙食費'],
    ['印刷費','6231文具用品'],
    ['雜費','6381雜項費用'],
  ];
  sh.getRange(1,1,rows.length,2).setValues(rows);
  sh.getRange(1,1,1,2).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/* ---------- 月分頁 ---------- */
function monthPrefix(){
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const roc = now.getFullYear() - 1911;
  const mm = Utilities.formatDate(now, tz, 'MM');
  return String(roc) + mm; // 例 11507
}

function getMonthSheet_(prefix){
  const ss = SpreadsheetApp.getActive();
  const name = MONTH_PREFIX + prefix;
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.getRange(2,COL_B,1,N_COLS).setValues([HEADER])
      .setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(2);
    sh.getRange('D:D').setNumberFormat('@');      // 日期存民國文字 115/07/05
    sh.getRange('H:H').setNumberFormat('#,##0');  // 金額千分位
    sh.setColumnWidth(7,320);                     // 摘要
    sh.getRange('G:G').setWrap(true);
  }
  return sh;
}

/* 合計列位置（I 欄 = '合計'），沒有則回 -1 */
function findSumRow_(sh){
  const last = sh.getLastRow();
  if(last < DATA_START_ROW) return -1;
  const vals = sh.getRange(DATA_START_ROW,9,last-DATA_START_ROW+1,1).getValues(); // I 欄
  for(let i=vals.length-1;i>=0;i--){
    if(String(vals[i][0]).trim() === '合計') return DATA_START_ROW+i;
  }
  return -1;
}

/* 該月最大 NO */
function maxNo_(sh){
  const last = sh.getLastRow();
  if(last < DATA_START_ROW) return 0;
  const vals = sh.getRange(DATA_START_ROW,COL_B,last-DATA_START_ROW+1,1).getValues();
  let m = 0;
  vals.forEach(r=>{
    const n = parseInt(r[0],10);
    if(!isNaN(n) && n > m) m = n;
  });
  return m;
}

function peekNextNo(){
  const prefix = monthPrefix();
  const sh = SpreadsheetApp.getActive().getSheetByName(MONTH_PREFIX + prefix);
  const next = (sh ? maxNo_(sh) : 0) + 1;
  return prefix + '-' + String(next).padStart(2,'0');
}

/* ---------- 儲存：直接寫成會計格式 ---------- */
function saveVoucher(body){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const prefix = monthPrefix();
    const sh = getMonthSheet_(prefix);
    const map = getSubjectMap_();
    const now = new Date();

    // NO：沿用前端帶來的（同月且未被別人用走），否則取月內下一號
    const maxNo = maxNo_(sh);
    let seq = 0;
    if(body.no){
      const parts = String(body.no).trim().split('-');
      if(parts.length === 2 && parts[0] === prefix) seq = parseInt(parts[1],10) || 0;
    }
    if(seq <= maxNo) seq = maxNo + 1;

    // 依科目分組（保持出現順序）：同科目併一列，不同科目各一列共用 NO
    const items = body.items || [];
    const order = [];
    const groups = {};
    items.forEach(it=>{
      const s = String(it.subject||'').trim();
      if(!groups[s]){ groups[s] = {memos:[], amounts:[]}; order.push(s); }
      if(it.memo) groups[s].memos.push(String(it.memo));
      groups[s].amounts.push(Number(it.amount)||0);
    });
    if(!order.length) return { ok:false, error:'沒有明細' };

    const rocDate = String(body.rocDate||'').replace(/-/g,'/'); // 115/07/05
    const rows = [];
    const unmappedRows = [];
    order.forEach((s,gi)=>{
      const g = groups[s];
      const acct = map[s] || '';
      if(!acct) unmappedRows.push(gi);
      const amount = g.amounts.length > 1 ? '=' + g.amounts.join('+') : (g.amounts[0]||0);
      rows.push([
        seq, '', "'" + rocDate, acct || s,
        '',                       // F 會計名稱：寫入後補 MID 公式
        g.memos.join('、'), amount,
        body.applicant||'', '', now  // 領款人欄記申請人姓名
      ]);
    });

    // 移除舊合計列 → 寫入資料 → 重加合計列
    const sumRow = findSumRow_(sh);
    if(sumRow > 0) sh.deleteRow(sumRow);
    const start = Math.max(DATA_START_ROW, sh.getLastRow()+1);
    sh.getRange(start, COL_B, rows.length, N_COLS).setValues(rows);
    rows.forEach((r,i)=>{
      const row = start + i;
      if(unmappedRows.indexOf(i) >= 0){
        sh.getRange(row,6).setValue(order[i]);                       // 未對照：F 直接放名稱
        sh.getRange(row,COL_B,1,N_COLS).setBackground('#fff2a8');    // 整列標黃提醒補對照
      }else{
        sh.getRange(row,6).setFormula('=MID(E'+row+',5,LEN(E'+row+'))');
      }
    });
    const lastData = start + rows.length - 1;
    sh.getRange(lastData+1, 8).setFormula('=SUM(H'+DATA_START_ROW+':H'+lastData+')');
    sh.getRange(lastData+1, 9).setValue('合計');
    sh.getRange(lastData+1, COL_B, 1, N_COLS).setFontWeight('bold');

    return { ok:true, no: prefix + '-' + String(seq).padStart(2,'0'), rows: rows.length,
             unmapped: unmappedRows.length };
  }finally{
    lock.releaseLock();
  }
}

/* ---------- 輔助表單：支出證明單 / 伙食費支出證明單 ----------
 * 單純流水帳（無特定格式），一張表單一列 */
function saveProof(b){
  const items = (b.items||[])
    .map(it => String(it.name||'') + (it.qty?('×'+it.qty):'') + (it.price?('@'+it.price):'') + (it.total?('='+it.total):''))
    .filter(s => s).join('；');
  return appendSimple_('支出證明單',
    ['建立時間','日期','科目','支出事由','不能取得單據之原因','受款者','明細','合計'],
    [new Date(), "'"+String(b.rocDate||'').replace(/-/g,'/'), b.subject||'', b.reason||'',
     b.why||'', b.payee||'', items, Number(b.total)||0]);
}

function saveFood(b){
  const meals = (b.meals||[])
    .map(m => [m.date,m.meal,m.people?(m.people+'人(桌)'):'',m.amount?('$'+m.amount):''].filter(String).join(' '))
    .filter(s => s).join('；');
  const buys = (b.buys||[])
    .map(x => String(x.name||'') + (x.qty?('×'+x.qty):'') + (x.amount?('='+x.amount):''))
    .filter(s => s).join('；');
  return appendSimple_('伙食費支出證明單',
    ['建立時間','日期','支出事由','期間自','期間至','受款者','人數','桌數','用餐情形','採買明細','附憑據張數','合計'],
    [new Date(), "'"+String(b.rocDate||'').replace(/-/g,'/'), b.reason||'',
     b.from||'', b.to||'', b.payee||'', b.people||'', b.tables||'',
     meals, buys, b.receipts||'', Number(b.total)||0]);
}

function appendSimple_(name, header, row){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if(!sh){
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.getRange(sh.getLastRow()+1, 1, 1, row.length).setValues([row]);
    return { ok:true };
  }finally{
    lock.releaseLock();
  }
}

/* ---------- 最近紀錄（預設本月，可帶 month=11506 查指定月） ---------- */
function listRecent(limit, month){
  const m = (month && /^\d{5}$/.test(String(month))) ? String(month) : monthPrefix();
  const sh = SpreadsheetApp.getActive().getSheetByName(MONTH_PREFIX + m);
  if(!sh) return [];
  const sumRow = findSumRow_(sh);
  const lastData = (sumRow > 0 ? sumRow - 1 : sh.getLastRow());
  if(lastData < DATA_START_ROW) return [];
  const start = Math.max(DATA_START_ROW, lastData - limit + 1);
  const values = sh.getRange(start, COL_B, lastData-start+1, N_COLS).getValues();
  return values.map(r=>{
    const o={}; HEADER.forEach((h,i)=> o[h]=r[i]); return o;
  });
}

/* ---------- 工具 ---------- */
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
