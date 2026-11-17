/**
 * 佳里奇美醫院論文網站 — 共用資料服務（準備版）
 *
 * Do NOT deploy this as "Anyone may edit".
 * No credentials or tokens belong in GitHub Pages or client JavaScript.
 *
 * Script Properties:
 *  SPREADSHEET_ID = destination Sheet ID
 *
 * Sheet tabs:
 *  Published: header row + raw records
 *  History: publishedAt, revision, rows, note
 *
 * This release is read-only over HTTP. Publishing must be done with the
 * owner's Google authorization via publishFromStaging(), after review.
 */
const HEADERS = [
  '流水號','申請序號','發表日期','申請者','申請者科別','申請項目',
  '論文類別','獎金','IF','學門排行','論文名稱','發表之雜誌/期刊',
  '發表地點','主管簽章','作者1','作者2','作者3','作者4','作者5','作者6'
];
const TAB_PUBLISHED='Published',TAB_STAGING='Staging',TAB_HISTORY='History';
function ss_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id)throw Error('Missing SPREADSHEET_ID script property');
  return SpreadsheetApp.openById(id);
}
function read_(tab) {
  const ws=ss_().getSheetByName(tab);
  if(!ws)return [];
  const rows=ws.getDataRange().getDisplayValues();
  if(!rows.length)return [];
  if(HEADERS.some((h,i)=>rows[0][i]!==h))throw Error(tab+' column mismatch');
  return rows.slice(1).filter(r=>r.some(v=>v!=='')).map(r=>r.slice(0,HEADERS.length));
}
function key_(r){return String(r[0]).trim()+'|'+String(r[1]).trim();}
function canonical_(r){
  const pua={'':'畊','':'彣','':'温','':'贋'};
  return JSON.stringify(r.map(v=>String(v??'').replace(/[-]/g,c=>pua[c]||c).replace(/\r\n/g,'\n')));
}
function validate_(candidate,existing){
  const seen=new Set(),old=new Map(existing.map(r=>[key_(r),canonical_(r)]));
  candidate.forEach((r,i)=>{
    if(!r[0]||!r[1]||!r[3])throw Error('Required application identity missing at row '+(i+2));
    const k=key_(r);if(seen.has(k))throw Error('Duplicate application at '+(i+2)+' '+k);
    seen.add(k);
    // Legacy publication-fee applications may have no title; preserve them,
    // but new applications must supply a title.
    if(!old.has(k)&&!r[10]&&!(r[0]==='1130117'&&r[1]==='202410A008'&&String(r[5]).includes('刊登費')))throw Error('New application title missing at row '+(i+2));
    if(old.has(k)&&old.get(k)!==canonical_(r))
      throw Error('Historical application modified: '+k);
  });
  const missing=[...old.keys()].filter(k=>!seen.has(k));
  if(missing.length)throw Error('Historical applications missing: '+missing.slice(0,5).join(', '));
  return {total:candidate.length,added:candidate.length-existing.length};
}
/** Read-only. Public sharing of the content must be approved by the hospital. */
function doGet(e) {
  try{
    const rows=read_(TAB_PUBLISHED);
    const result={ok:true,version:1,updatedAt:updatedAt_(),headers:HEADERS,rows};
    const callback=e&&e.parameter&&e.parameter.callback;
    if(callback){
      // Safe identifier only; do not interpolate arbitrary JavaScript.
      if(!/^[A-Za-z_$][\w$]{0,70}$/.test(callback))throw Error('Invalid callback');
      return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function updatedAt_(){
  const ws=ss_().getSheetByName(TAB_HISTORY);
  if(!ws||ws.getLastRow()<2)return '';
  return String(ws.getRange(ws.getLastRow(),1).getDisplayValue());
}
/**
 * Explicitly disabled: an unauthenticated browser POST must not be able to
 * overwrite a hospital-wide dataset.
 */
function doPost(){
  return ContentService.createTextOutput(JSON.stringify({
    ok:false,error:'Public writes disabled. Use approved publish workflow.'
  })).setMimeType(ContentService.MimeType.JSON);
}
/**
 * Publish approved Staging data. The operator must open the Apps Script editor
 * and authorize the operation with their Google account.
 */
function publishFromStaging(){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const ss=ss_(),candidate=read_(TAB_STAGING),existing=read_(TAB_PUBLISHED);
    if(!candidate.length)throw Error('Staging is empty');
    const checked=validate_(candidate,existing);
    const published=ss.getSheetByName(TAB_PUBLISHED)||ss.insertSheet(TAB_PUBLISHED);
    const history=ss.getSheetByName(TAB_HISTORY)||ss.insertSheet(TAB_HISTORY);
    const revision=Utilities.getUuid(),at=new Date().toISOString();
    // Backup is a full sheet copy, not just a summary row.
    if(existing.length){
      const backup=ss.insertSheet('Backup_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'MMdd_HHmmss'));
      backup.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
      backup.getRange(2,1,existing.length,HEADERS.length).setValues(existing);
      backup.hideSheet();
    }
    published.clearContents();
    published.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    published.getRange(2,1,candidate.length,HEADERS.length).setNumberFormat('@').setValues(candidate);
    if(history.getLastRow()===0)history.appendRow(['publishedAt','revision','rows','note']);
    history.appendRow([at,revision,candidate.length,'新增 '+checked.added+' 筆；已保留歷史紀錄']);
    SpreadsheetApp.flush();
    return {ok:true,revision,total:checked.total,added:checked.added};
  }finally{lock.releaseLock();}
}
