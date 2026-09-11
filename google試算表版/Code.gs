/* IP 週報 App 的後端：綁在 Google 試算表上的 Apps Script
 * 部署方式見同資料夾「部署步驟.md」。App 用 GET 讀、POST 寫。
 */
const R_HEAD = ['week','editor','ip','status','pending','chase','news','pass','assets','assetsNote','help','opening','updatedAt','film','stage','reason','reasonNote','ipReply','replyNote','chased','reviewDate','publishDate','photos','links'];
const E_HEAD = ['editor','ips','updatedAt'];

function sheetOf(name, head){
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if(!s){ s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); }
  const h = s.getRange(1,1,1,head.length).getValues()[0]; if(h.join('\u0001')!==head.join('\u0001')) s.getRange(1,1,1,head.length).setValues([head]);   // 欄位有新增時補標題
  s.getRange('A:Z').setNumberFormat('@');   // 全部當文字，避免日期自動轉換
  return s;
}
/* 試算表會把 2026-09-11 之類的字自動變成日期，讀回來一律轉回文字（週＝yyyy-MM-dd，時間＝ISO） */
function tz(){ return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
function txtDay(v){ if(v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, tz(), 'yyyy-MM-dd'); return String(v==null?'':v); }
function txtIso(v){ if(v instanceof Date) return isNaN(v) ? '' : v.toISOString(); return String(v==null?'':v); }
function txt(v){ if(v instanceof Date) return txtIso(v); return String(v==null?'':v); }
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.action === 'ping') return out({ ok:true, app:'ip-report', time:new Date().toISOString() });
  if(p.action === 'all')  return out(getAll(p.week || ''));
  return out({ ok:false, error:'unknown action' });
}

function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents || '{}'); }catch(err){ return out({ ok:false, error:'bad json' }); }
  if(body.action === 'upload') return out(uploadPhoto(body));   // 上傳截圖不用排隊
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    if(body.action === 'save') return out(saveEntry(body));
    if(body.action === 'ips')  return out(setIps(body.editor, body.ips));
    return out({ ok:false, error:'unknown action' });
  } finally { lock.releaseLock(); }
}

/* 讀：所有人的回報（weeks[日期][IP] = 影片陣列）+ 每個人的 IP 清單 */
function getAll(week){
  const rs = sheetOf('回報', R_HEAD), es = sheetOf('剪輯師', E_HEAD);
  const editors = {};
  const ev = es.getDataRange().getValues();
  for(let i=1;i<ev.length;i++){ const [editor, ips, updatedAt] = ev[i]; if(!editor) continue;
    editors[editor] = { name:txt(editor), ips: safeJson(ips, []), updatedAt: txt(updatedAt), weeks:{} }; }
  const rv = rs.getDataRange().getValues();
  for(let i=1;i<rv.length;i++){ const row = rv[i]; const o = {}; R_HEAD.forEach((k,j)=>o[k]=txt(row[j])); o.week = txtDay(row[0]);
    if(!o.editor || !o.ip) continue; if(week && o.week!==week) continue;
    const ed = editors[o.editor] || (editors[o.editor] = { name:String(o.editor), ips:[], updatedAt:'', weeks:{} });
    const wk = ed.weeks[o.week] || (ed.weeks[o.week] = {});
    (wk[o.ip] = wk[o.ip] || []).push({ film:String(o.film||''), stage:String(o.stage||''), reason:String(o.reason||''), reasonNote:String(o.reasonNote||''),
      ipReply:String(o.ipReply||''), replyNote:String(o.replyNote||''), pending:String(o.pending||''), chased:String(o.chased||''),
      reviewDate:txtDay(o.reviewDate), publishDate:txtDay(o.publishDate), news:String(o.news||''), help:String(o.help||''),
      photos:safeJson(o.photos, []), links:safeJson(o.links, []),
      status:String(o.status||''), updatedAt:String(o.updatedAt||'') });   // 一個 IP 一天可以有很多支片
  }
  return { ok:true, week, editors };
}

/* 寫：一個人、一天、一個 IP 的一支片（同片名覆蓋；keyFilm＝原片名，改名時用） */
function saveEntry(b){
  const rs = sheetOf('回報', R_HEAD);
  const e = b.entry || {}; const key = [String(b.week), String(b.editor), String(b.ip)];
  const keyFilm = (b.keyFilm !== undefined && b.keyFilm !== null) ? String(b.keyFilm) : String(e.film||'');   // 用「日期＋人＋IP＋片名」找同一支
  const vals = rs.getDataRange().getValues();
  let rowIdx = -1;
  for(let i=1;i<vals.length;i++){ if(txtDay(vals[i][0])===key[0] && txt(vals[i][1])===key[1] && txt(vals[i][2])===key[2] && txt(vals[i][13])===keyFilm){ rowIdx=i+1; break; } }
  const now = new Date().toISOString();
  const row = [key[0], key[1], key[2], e.status||'', e.pending||'', e.chase||'', e.news||'', e.pass||'無', e.assets||'無', e.assetsNote||'', e.help||'', JSON.stringify(e.opening||[]), now, e.film||'', e.stage||'', e.reason||'', e.reasonNote||'', e.ipReply||'', e.replyNote||'', e.chased||'', e.reviewDate||'', e.publishDate||'', JSON.stringify(e.photos||[]), JSON.stringify(e.links||[])];
  if(rowIdx>0){
    const old = txt(vals[rowIdx-1][12]);
    if(e.updatedAt && old && old > e.updatedAt && b.force!==true) return { ok:true, skipped:true, reason:'server newer', updatedAt: old };
    rs.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else rs.appendRow(row);
  if(Array.isArray(b.ips)) setIps(b.editor, b.ips);
  return { ok:true, updatedAt: now };
}

function setIps(editor, ips){
  const es = sheetOf('剪輯師', E_HEAD);
  const vals = es.getDataRange().getValues(); const now = new Date().toISOString();
  for(let i=1;i<vals.length;i++){ if(txt(vals[i][0])===String(editor)){ es.getRange(i+1,1,1,3).setValues([[editor, JSON.stringify(ips||[]), now]]); return { ok:true }; } }
  es.appendRow([editor, JSON.stringify(ips||[]), now]); return { ok:true };
}

/* 截圖：存到雲端硬碟資料夾「IP週報照片」，開「知道連結的人可檢視」，回傳可直接當 <img> 的網址 */
function photoFolder(){ const it = DriveApp.getFoldersByName('IP週報照片'); return it.hasNext() ? it.next() : DriveApp.createFolder('IP週報照片'); }
function uploadPhoto(b){
  if(!b.data) return { ok:false, error:'沒有圖片' };
  const bytes = Utilities.base64Decode(b.data);
  if(bytes.length > 8*1024*1024) return { ok:false, error:'圖片太大（超過 8MB）' };
  const blob = Utilities.newBlob(bytes, b.mime || 'image/jpeg', b.name || ('截圖_' + new Date().toISOString().replace(/[:.]/g,'-') + '.jpg'));
  const f = photoFolder().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, id:f.getId(), url:'https://lh3.googleusercontent.com/d/' + f.getId(), link:f.getUrl() };
}

function safeJson(v, def){ try{ return v ? JSON.parse(v) : def; }catch(e){ return def; } }
