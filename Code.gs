/*************************************************************
 * 黎明教會 付款申請單（代傳票）後端 API + 會計月報
 * 部署：部署 → 新增部署作業 → 網頁應用程式
 *   執行身分：我
 *   誰可以存取：任何人
 * 把 /exec 網址貼到前端 index.html 的 GAS_API_URL
 *
 * Google Sheet 分頁：
 *   1.「科目」   A 欄科目名稱（前端下拉用）、B 欄會計科目（代號+名稱，產月報用）
 *               可從選單「黎明教會 → 初始化科目分頁」建立預設內容
 *   2.「申請單」 流水帳，由程式自動建立標題列
 *   3.「會計支出115MM」 月報分頁，由選單「產生會計月報」自動生成
 *************************************************************/

const SHEET_SUBJECT = '科目';
const SHEET_DATA    = '申請單';
const REPORT_PREFIX = '會計支出';
const HEADERS = ['流水號','序','日期','民國日期','科目','摘要','金額','領款人','申請人','合計','建立時間','備註'];

/* ---------- 試算表選單 ---------- */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('黎明教會')
    .addItem('產生本月會計月報','reportThisMonth')
    .addItem('產生指定月份月報…','reportPickMonth')
    .addSeparator()
    .addItem('初始化科目分頁','initSubjectSheet')
    .addToUi();
}

/* ---------- 路由 ---------- */
function doGet(e){
  const action = (e.parameter.action || '').toLowerCase();
  if(action === 'subjects') return json(getSubjects());
  if(action === 'nextno')   return json({ no: peekNextNo() });
  if(action === 'list')     return json(listRecent(Number(e.parameter.limit) || 20));
  return json({ ok:true, msg:'黎明教會 付款申請單 API' });
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.action === 'save') return json(saveVoucher(body));
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

/* ---------- 流水號：民國年月 + 月內序，如 11507-16 ----------
 * 同一張申請單多列共用同一號；月報 NO 直接取「-」後的序號 */
function monthPrefix(){
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const roc = now.getFullYear() - 1911;
  const mm = Utilities.formatDate(now, tz, 'MM');
  return String(roc) + mm; // 例 11507
}
function peekNextNo(){ return buildNo_(); }
function buildNo_(){
  const sh = getDataSheet_();
  const prefix = monthPrefix();
  const last = sh.getLastRow();
  let maxSeq = 0;
  if(last >= 2){
    const nos = sh.getRange(2,1,last-1,1).getValues();
    nos.forEach(r=>{
      const s = String(r[0]);
      if(s.indexOf(prefix + '-') === 0){
        const seq = parseInt(s.split('-')[1],10);
        if(!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }
  return prefix + '-' + String(maxSeq + 1).padStart(2,'0');
}

/* ---------- 儲存（一張單多列，共用流水號） ---------- */
function saveVoucher(body){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const sh  = getDataSheet_();
    const no  = body.no && String(body.no).trim() ? String(body.no).trim() : buildNo_();
    const now = new Date();
    const items = body.items || [];
    // 民國日期加 ' 前綴強制存為文字，避免 Sheet 誤判成西元 115 年
    const roc = body.rocDate ? "'" + body.rocDate : '';
    const subject = String(body.subject||'').trim(); // 一張單一個科目
    const rows = items.map((it,i)=>([
      no, i+1, body.date||'', roc,
      it.subject||subject, it.memo||'', Number(it.amount)||0,
      body.payee||'', body.applicant||'', Number(body.total)||0, now, ''
    ]));
    if(rows.length){
      sh.getRange(sh.getLastRow()+1, 1, rows.length, HEADERS.length).setValues(rows);
    }
    return { ok:true, no:no, rows:rows.length };
  }finally{
    lock.releaseLock();
  }
}

/* ---------- 會計月報：彙整當月申請單成「會計支出115MM」分頁 ---------- */
function reportThisMonth(){ buildMonthlyReport_(monthPrefix()); }

function reportPickMonth(){
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('產生月報','輸入民國年月（5 碼，例 11507）：',ui.ButtonSet.OK_CANCEL);
  if(res.getSelectedButton() !== ui.Button.OK) return;
  const m = res.getResponseText().trim();
  if(!/^\d{5}$/.test(m)){ ui.alert('格式錯誤，請輸入 5 碼民國年月，例 11507'); return; }
  buildMonthlyReport_(m);
}

function buildMonthlyReport_(month){
  const ss = SpreadsheetApp.getActive();
  const data = getDataSheet_();
  const map = getSubjectMap_();
  const last = data.getLastRow();

  // 同一張單（流水號）同一科目的多列明細，併成月報一列：
  // 摘要用「、」串接，金額寫成 =a+b+c 公式（比照會計現行格式）
  const groups = {};
  const order = [];
  if(last >= 2){
    data.getRange(2,1,last-1,HEADERS.length).getValues().forEach(r=>{
      const no = String(r[0]);
      if(no.indexOf(month + '-') !== 0) return;
      const seq = parseInt(no.split('-')[1],10) || 0;
      const subject = String(r[4]||'').trim();
      const key = seq + '|' + subject;
      if(!groups[key]){
        groups[key] = {
          seq: seq, line: Number(r[1])||0,
          rocDate: String(r[3]||'').replace(/-/g,'/'),
          subject: subject, memos: [], amounts: [],
          payee: r[7], note: r[11]||''
        };
        order.push(key);
      }
      const g = groups[key];
      if(r[5]) g.memos.push(String(r[5]));
      g.amounts.push(Number(r[6])||0);
    });
  }
  const rows = order.map(k=>groups[k]);
  rows.sort((a,b)=> a.seq - b.seq || a.line - b.line);

  const name = REPORT_PREFIX + month;
  let sh = ss.getSheetByName(name);
  if(sh) sh.clear(); else sh = ss.insertSheet(name);

  // 版面比照會計 xlsx：B2 起，B=NO C=月 D=日期 E=科目代號 F=會計名稱 G=摘要 H=合計 I=領款人 J=備註
  sh.getRange(2,2,1,9).setValues([['NO','月','日期','科目代號','會計名稱','摘要','合計','領款人','備註']])
    .setFontWeight('bold').setHorizontalAlignment('center');

  const unmapped = [];
  if(rows.length){
    const values = rows.map(r=>{
      const acct = map[r.subject] || '';
      if(!acct) unmapped.push(r);
      const memo = r.memos.join('、');
      // 多筆明細 → =a+b+c 公式；單筆 → 直接填數字
      const amount = r.amounts.length > 1 ? '=' + r.amounts.join('+') : (r.amounts[0]||0);
      return [r.seq, '', r.rocDate, acct || r.subject, '', memo, amount, r.payee, r.note];
    });
    sh.getRange(3,2,values.length,9).setValues(values);
    // 會計名稱：有對照的用 MID 公式切出名稱；未對照的整列標黃提醒
    rows.forEach((r,i)=>{
      const row = 3 + i;
      if(map[r.subject]){
        sh.getRange(row,6).setFormula('=MID(E'+row+',5,LEN(E'+row+'))');
      }else{
        sh.getRange(row,6).setValue(r.subject);
        sh.getRange(row,2,1,9).setBackground('#fff2a8');
      }
    });
  }

  const sumRow = 3 + rows.length;
  sh.getRange(sumRow,8).setFormula(rows.length ? '=SUM(H3:H'+(sumRow-1)+')' : '=0');
  sh.getRange(sumRow,9).setValue('合計');
  sh.getRange(sumRow,2,1,9).setFontWeight('bold');

  sh.getRange(3,8,Math.max(rows.length,1)+1,1).setNumberFormat('#,##0');
  sh.setColumnWidth(7,320);  // 摘要
  sh.getRange(3,7,Math.max(rows.length,1),1).setWrap(true);
  [2,3,4,6,8,9,10].forEach(c=> sh.autoResizeColumn(c));

  const ui = SpreadsheetApp.getUi();
  ui.alert('月報「'+name+'」已產生：'+rows.length+' 列' +
    (unmapped.length ? '\n\n注意：有 '+unmapped.length+' 列科目未對照到會計科目（已標黃），請到「科目」分頁 B 欄補上對照後重新產生。' : ''));
}

/* ---------- 最近紀錄（可選，給日後查詢頁用） ---------- */
function listRecent(limit){
  const sh = getDataSheet_();
  const last = sh.getLastRow();
  if(last < 2) return [];
  const start = Math.max(2, last - limit + 1);
  const values = sh.getRange(start,1,last-start+1,HEADERS.length).getValues();
  return values.map(r=>{
    const o={}; HEADERS.forEach((h,i)=> o[h]=r[i]); return o;
  });
}

/* ---------- 工具 ---------- */
function getDataSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_DATA);
  if(!sh){
    sh = ss.insertSheet(SHEET_DATA);
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('D2:D').setNumberFormat('@'); // 民國日期欄固定為文字格式
  }
  return sh;
}
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
