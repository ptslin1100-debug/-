/**
 * 班級分組後端 — Google Sheets + Google Apps Script
 * 1. 在 script.google.com 建立獨立專案，貼上本檔內容。
 * 2. 在編輯器選「setup」並執行一次，允許權限。
 * 3. 到「執行紀錄」取得新建立的 Google Sheet 連結；
 *    在「個人代碼」分頁私下發放各同學代碼，絕對不要公開這個分頁。
 * 4. 部署 > 新增部署 > 網頁應用程式 > 執行身分：我 > 誰可以存取：任何人。
 * 5. 將部署取得的 /exec 網址填入 GitHub 的 config.js。
 *
 * 安全說明：以每人唯一代碼驗證寫入，代碼不進 GitHub。
 * 公開的讀取 API 只傳名單與分組狀態；不得公開試算表本體。
 */
const ROSTER = [["01","江榮展","Nick"],["02","廖煒立","Willy"],["03","徐其祥","Mark"],["04","於欣慧","Kelly"],["05","洪靖宜","Amber"],["06","翁加偉","William"],["07","曾茂榮",""],["08","陳韋伊","Wendy"],["09","李思萱","Alicia"],["11","林逸賢","James"],["12","高硯清","Maria"],["14","董建一","Bryan"],["15","呂逸庭","Grace"],["16","王鈞賢","Jason"],["17","陳柔安","Ruby"],["18","馬瑞琪","Floria"],["19","陳宥彤",""],["20","李佳樺","Cathy"],["21","彭思嘉","Scarlett"],["22","張志騰","Rick"],["23","林佳樺",""],["24","連丈賢","Daniel"],["26","曾雯瑾",""],["27","陳昀資",""],["28","鄧棋耀","Cassell"],["29","趙克秦","Iris"],["30","張稚婉","Phoenix"],["31","石在元","Barry"],["32","余佩容","Judy"],["33","黃詩慧","Sylvie"],["34","朱泙慧","Claire"],["35","賴佳妤","Casey"],["36","吳文瑜","Amanda"],["37","葉日如","SnowGuy"],["38","楊芷芸","Stella"],["39","盧怡伶","Bonnie"],["40","余禮安","Lian"]];
const SHEET_NAME='分組狀態', CODE_SHEET='個人代碼', GROUP_SHEET='組別名稱';
function setup() {
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('SHEET_ID')) {
    Logger.log('已初始化。Sheet: https://docs.google.com/spreadsheets/d/'+props.getProperty('SHEET_ID')+'/edit');
    return;
  }
  const ss=SpreadsheetApp.create('班級分組共用管理表（37人）');
  const sheet=ss.getSheets()[0]; sheet.setName(SHEET_NAME);
  sheet.getRange(1,1,1,6).setValues([['座號','姓名','綽號','狀態','組別','更新時間']]);
  sheet.getRange(2,1,ROSTER.length,6).setValues(ROSTER.map(p=>[...p,'尋找組別','',new Date()]));
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#e1f0e7');
  sheet.autoResizeColumns(1,6);
  const codes=ss.insertSheet(CODE_SHEET);
  codes.getRange(1,1,1,3).setValues([['座號','姓名','個人專屬代碼（請私下發送）']]);
  codes.getRange(2,1,ROSTER.length,3).setValues(ROSTER.map(p=>[p[0],p[1],Utilities.getUuid().replace(/-/g,'').slice(0,12)]));
  codes.getRange(1,1,1,3).setFontWeight('bold').setBackground('#ffe8ba');
  codes.setFrozenRows(1); codes.autoResizeColumns(1,3);
  const names=ss.insertSheet(GROUP_SHEET);
  names.getRange(1,1,1,2).setValues([['組別','組名']]);
  names.getRange(2,1,7,2).setValues(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組']));
  names.setFrozenRows(1); names.autoResizeColumns(1,2);
  const names=ss.insertSheet(GROUP_SHEET); names.getRange(1,1,1,2).setValues([['組別','組名']]); names.getRange(2,1,7,2).setValues(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組']));
  props.setProperty('SHEET_ID',ss.getId());
  Logger.log('建立成功，Google Sheet: '+ss.getUrl());
}
function getBook_(){
  const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if(!id) throw Error('請先在 Apps Script 執行 setup()');
  return SpreadsheetApp.openById(id);
}
function publicRows_(){
  const sh=getBook_().getSheetByName(SHEET_NAME);
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues().map(r=>({
    seat:String(r[0]).padStart(2,'0'),name:r[1],nickname:r[2],
    status:r[3]||'尋找組別',group_no:r[4]===''?null:Number(r[4])
  }));
}
function groupNames_(){
  const sh=getBook_().getSheetByName(GROUP_SHEET);
  if(!sh)return Object.fromEntries(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組']));
  return Object.fromEntries(sh.getRange(2,1,7,2).getValues().map(r=>[Number(r[0]),String(r[1]||'第 '+r[0]+' 組')]));
}
function authorize_(ss,seat,code){
  const sh=ss.getSheetByName(CODE_SHEET), values=sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues();
  const found=values.find(x=>String(x[0]).padStart(2,'0')===seat);
  if(!found||!code||found[2]!==code)throw Error('座號或個人代碼不正確');
}
function renameGroup_(p){
  const seat=String(p.seat||'').padStart(2,'0'), code=String(p.code||'').trim();
  const group=Number(p.group_no), name=String(p.group_name||'').trim();
  if(!Number.isInteger(group)||group<1||group>7)throw Error('組別不正確');
  if(name.length<1||name.length>24)throw Error('組名長度必須為 1–24 字');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const ss=getBook_();authorize_(ss,seat,code);
    const mine=publicRows_().find(r=>r.seat===seat);
    if(!mine||mine.group_no!==group||!['確定入組','考慮中'].includes(mine.status))throw Error('只能修改自己所在或考慮中的組別名稱');
    let sh=ss.getSheetByName(GROUP_SHEET);
    if(!sh){sh=ss.insertSheet(GROUP_SHEET);sh.getRange(1,1,1,2).setValues([['組別','組名']]);sh.getRange(2,1,7,2).setValues(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組']));}
    sh.getRange(group+1,2).setValue(name);SpreadsheetApp.flush();
    return {ok:true};
  } finally {lock.releaseLock();}
}
function groupNames_(){const sh=getBook_().getSheetByName(GROUP_SHEET); if(!sh)return Object.fromEntries(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組'])); return Object.fromEntries(sh.getRange(2,1,7,2).getValues().map(r=>[Number(r[0]),String(r[1]||'第 '+r[0]+' 組')]));}
function authorize_(ss,seat,code){const sh=ss.getSheetByName(CODE_SHEET);const values=sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues(); const stored=values.find(x=>String(x[0]).padStart(2,'0')===seat);if(!stored||!code||stored[2]!==code)throw Error('座號或個人代碼不正確')}
function renameGroup_(p){
const seat=String(p.seat||'').padStart(2,'0'),code=String(p.code||'').trim(),g=Number(p.group_no),name=String(p.group_name||'').trim();
if(!Number.isInteger(g)||g<1||g>7)throw Error('組別不正確');
if(!name||name.length>24)throw Error('組名須為 1–24 字');
const lock=LockService.getScriptLock();lock.waitLock(10000);
try{const ss=getBook_();authorize_(ss,seat,code);const mine=publicRows_().find(r=>r.seat===seat);if(!mine||mine.group_no!==g||mine.status!=='確定入組')throw Error('只有此組的正式組員能修改組名');
let sh=ss.getSheetByName(GROUP_SHEET);if(!sh){sh=ss.insertSheet(GROUP_SHEET);sh.getRange(1,1,1,2).setValues([['組別','組名']]);sh.getRange(2,1,7,2).setValues(Array.from({length:7},(_,i)=>[i+1,'第 '+(i+1)+' 組']));}
sh.getRange(g+1,2).setValue(name);SpreadsheetApp.flush();return {ok:true};
}finally{lock.releaseLock();}
}
function doGet(e){
  const callback=String(e.parameter.callback||'');
  if(!/^[a-zA-Z_$][\w$\.]{0,99}$/.test(callback)) return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
  let data;
  try {data={ok:true,roster:publicRows_(),groupNames:groupNames_(),at:new Date().toISOString()};}
  catch(err){data={ok:false,error:String(err.message)};}
  return ContentService.createTextOutput(callback+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function doPost(e){
  let result;
  try {const p=e.parameter||{};result=p.action==='rename'?renameGroup_(p):save_(p);}
  catch(err){result={ok:false,error:String(err.message)};}
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
function save_(p){
  const seat=String(p.seat||'').padStart(2,'0');
  const code=String(p.code||'').trim();
  const status=String(p.status||'');
  const valid=['確定入組','考慮中','尋找組別'];
  if(!valid.includes(status)) throw Error('狀態不正確');
  const g=status==='尋找組別'?null:Number(p.group_no);
  if(status!=='尋找組別'&&(!Number.isInteger(g)||g<1||g>7)) throw Error('組別不正確');
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const ss=getBook_(), sheet=ss.getSheetByName(SHEET_NAME), codeSheet=ss.getSheetByName(CODE_SHEET);
    authorize_(ss,seat,code);
    const roster=publicRows_();
    const idx=roster.findIndex(r=>r.seat===seat);
    if(idx<0) throw Error('不存在此座號');
    if(status==='確定入組'){
      const confirmed=roster.filter(r=>r.seat!==seat&&r.status==='確定入組'&&r.group_no===g).length;
      if(confirmed >= (g===7?7:6)) throw Error('這組正式名額已滿');
    }
    const row=idx+2;
    sheet.getRange(row,4,1,3).setValues([[status,g===null?'':g,new Date()]]);
    SpreadsheetApp.flush();
    return {ok:true};
  } finally {lock.releaseLock();}
}
