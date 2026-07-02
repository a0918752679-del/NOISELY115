
const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });

const PORT = Number(process.env.PORT || 8080);
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID || '';
const GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '';
const GOOGLE_AUTO_SYNC = String(process.env.GOOGLE_AUTO_SYNC || 'false').toLowerCase() === 'true';
const SHEET_RAW = process.env.GOOGLE_SHEET_RAW_NAME || '執行成效資料';
const SHEET_VEHICLES = process.env.GOOGLE_SHEET_VEHICLE_NAME || '車輛案件資料';
const GOOGLE_SHEET_URL = process.env.GOOGLE_SHEET_URL || (GOOGLE_SHEET_ID ? `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/edit` : 'https://docs.google.com/spreadsheets/d/1EfP7GoI87RRl1AUGegwPhqNvFN9xG-YXm9NoMSqm_O0/edit?gid=0#gid=0');
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN || '';
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET || '';
const LINE_WEBHOOK_VERIFY = String(process.env.LINE_WEBHOOK_VERIFY || 'false').toLowerCase() === 'true';
const DASHBOARD_ADMIN_TOKEN = process.env.DASHBOARD_ADMIN_TOKEN || process.env.ADMIN_TOKEN || '';
const DASHBOARD_EDITOR_TOKEN = process.env.DASHBOARD_EDITOR_TOKEN || '';

const DATA_FILE = path.join(__dirname, 'data', 'dashboard-cache.json');
const BUILTIN_DATA_FILE = path.join(__dirname, 'data', 'builtin-dashboard-cache.json');
const RAW_COLUMNS = ['場次編號','機台編號','案件來源','執行時段','日期','月份','行政區','點位地址','辨識車流','超標數','告發件數','通知到檢件數','告發金額','是否完成'];
const VEHICLE_COLUMNS = ['案件類型','車牌','車種','日期','量測時間','行政區','點位地址','道路','量測值','標準值','超標值','金額','案件編號','官方註記','來源備註'];

function ensureDataDir(){ fs.mkdirSync(path.dirname(DATA_FILE), { recursive: true }); }
function ensureData(){
  ensureDataDir();
  if(!fs.existsSync(DATA_FILE)){
    if(fs.existsSync(BUILTIN_DATA_FILE)){
      fs.copyFileSync(BUILTIN_DATA_FILE, DATA_FILE);
    }else{
      fs.writeFileSync(DATA_FILE, JSON.stringify({ raw: [], vehicles: [], updatedAt: '' }, null, 2), 'utf8');
    }
  }
}
function readData(){ ensureData(); try{return JSON.parse(fs.readFileSync(DATA_FILE,'utf8'))}catch{return {raw:[],vehicles:[],updatedAt:''}} }
function writeData(data){ ensureDataDir(); fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8'); }
function googleSheetConfigured(){ return Boolean(GOOGLE_SHEET_ID && GOOGLE_SERVICE_ACCOUNT_JSON); }
function b64url(input){ return Buffer.from(input).toString('base64').replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_'); }
function getGoogleCredentials(){
  if(!GOOGLE_SERVICE_ACCOUNT_JSON) throw new Error('未設定 GOOGLE_SERVICE_ACCOUNT_JSON。');
  const raw = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  try{ return raw.startsWith('{') ? JSON.parse(raw) : JSON.parse(Buffer.from(raw,'base64').toString('utf8')); }
  catch(e){ throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON 格式錯誤，請貼 JSON 原文或 base64。'); }
}
async function getGoogleAccessToken(){
  const c = getGoogleCredentials();
  const now = Math.floor(Date.now()/1000);
  const header = {alg:'RS256',typ:'JWT'};
  const claim = {iss:c.client_email,scope:'https://www.googleapis.com/auth/spreadsheets',aud:'https://oauth2.googleapis.com/token',exp:now+3600,iat:now};
  const unsigned = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const signer = crypto.createSign('RSA-SHA256'); signer.update(unsigned);
  const sig = signer.sign(c.private_key,'base64').replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
  const res = await fetch('https://oauth2.googleapis.com/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({grant_type:'urn:ietf:params:oauth:grant-type:jwt-bearer',assertion:`${unsigned}.${sig}`})});
  const json = await res.json().catch(()=>({}));
  if(!res.ok) throw new Error(`Google Token 取得失敗：${json.error_description||json.error||res.status}`);
  return json.access_token;
}
function encRange(range){ return encodeURIComponent(range).replace(/%21/g,'!'); }
async function googleSheetsRequest(method, url, body){
  const token = await getGoogleAccessToken();
  const res = await fetch(url,{method,headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:body?JSON.stringify(body):undefined});
  const text = await res.text(); let json={}; try{json=text?JSON.parse(text):{}}catch{json={raw:text}}
  if(!res.ok) throw new Error(json?.error?.message || text || `Google Sheets API ${res.status}`);
  return json;
}
async function ensureNamedSheetExists(name){
  if(!GOOGLE_SHEET_ID) throw new Error('未設定 GOOGLE_SHEET_ID。');
  const meta = await googleSheetsRequest('GET', `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}`);
  const exists = (meta.sheets||[]).some(s => s.properties?.title === name);
  if(!exists){
    await googleSheetsRequest('POST', `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}:batchUpdate`, {requests:[{addSheet:{properties:{title:name}}}]});
  }
}
function normalizeHeader(v){ return String(v ?? '').trim().toLowerCase().replace(/[\s_\-／/（）()：:]/g,''); }
function valueBy(row,aliases){
  const normalized={}; Object.entries(row||{}).forEach(([k,v])=>normalized[normalizeHeader(k)] = v);
  for(const a of aliases){ const k=normalizeHeader(a); if(Object.prototype.hasOwnProperty.call(normalized,k)) return normalized[k]; }
  return '';
}
function toNum(v){ if(v===null||v===undefined||v==='') return 0; const n=Number(String(v).replace(/,/g,'').replace(/元|件|場|db/ig,'')); return Number.isFinite(n)?n:0; }
function toBool(v){ if(typeof v==='boolean')return v; const t=String(v??'').trim().toLowerCase(); return !['false','0','否','未完成','no','n'].includes(t); }
function pad2(v){ return String(v).padStart(2,'0'); }
function excelSerialToDate(n){ const d=new Date(Math.round((Number(n)-25569)*86400*1000)); return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth()+1)}-${pad2(d.getUTCDate())}`; }
function normalizeDate(v){
  if(v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad2(v.getMonth()+1)}-${pad2(v.getDate())}`;
  if(typeof v==='number' && v>20000) return excelSerialToDate(v);
  let t=String(v??'').trim(); if(!t)return '';
  if(/^\d+(\.\d+)?$/.test(t) && Number(t)>20000) return excelSerialToDate(Number(t));
  t=t.replace(/[/.]/g,'-').replace(/年/g,'-').replace(/月/g,'-').replace(/日/g,'').replace(/\s+.*/,'');
  const m=t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); if(m)return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  const roc=t.match(/^(\d{2,3})-(\d{1,2})-(\d{1,2})$/); if(roc)return `${Number(roc[1])+1911}-${pad2(roc[2])}-${pad2(roc[3])}`;
  return t;
}
function normalizeDateTime(v,dateFallback=''){
  if(v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad2(v.getMonth()+1)}-${pad2(v.getDate())} ${pad2(v.getHours())}:${pad2(v.getMinutes())}:${pad2(v.getSeconds())}`;
  let t=String(v??'').trim(); if(!t)return dateFallback;
  t=t.replace(/[年月]/g,'-').replace(/日/g,'').replace(/\//g,'-');
  const compact=t.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}:\d{2}:\d{2})$/); if(compact)return `${compact[1]}-${compact[2]}-${compact[3]} ${compact[4]}`;
  const roc=t.match(/^(\d{3})-(\d{2})-(\d{2})\s+(\d{1,2}:\d{2}:\d{2})$/); if(roc)return `${Number(roc[1])+1911}-${roc[2]}-${roc[3]} ${roc[4]}`;
  return t;
}
function extractDistrict(location){ const m=String(location||'').match(/(萬里|金山|板橋|汐止|深坑|石碇|瑞芳|平溪|雙溪|貢寮|新店|坪林|烏來|永和|中和|土城|三峽|樹林|鶯歌|三重|新莊|泰山|林口|蘆洲|五股|八里|淡水|三芝|石門)區/); return m?m[0]:''; }
function stripDistrict(location,district){ let s=String(location||'').replace(/^新北市/,'').trim(); if(district&&s.startsWith(district))s=s.slice(district.length); return s.trim(); }
function mapRawRow(row,index=0){
  const date=normalizeDate(valueBy(row,['日期','date','執行日期','監測日期']));
  const seqRaw=valueBy(row,['場次編號','場次','序號','seq','id']);
  return {
    seq:toNum(seqRaw)||index+1,
    machine:String(valueBy(row,['機台編號','機台','machine','設備編號'])||'').trim(),
    caseType:String(valueBy(row,['案件來源','案件類型','caseType','來源'])||'').trim(),
    period:String(valueBy(row,['執行時段','時段','period'])||'').trim(),
    date,
    month:toNum(valueBy(row,['月份','month']))||toNum((date.split('-')[1]||'')),
    district:String(valueBy(row,['行政區','district','區域'])||'').trim(),
    location:String(valueBy(row,['點位地址','點位','地址','location','監測地點'])||'').trim(),
    vehicles:toNum(valueBy(row,['辨識車流','車流辨識量','車流','vehicles'])),
    over:toNum(valueBy(row,['超標數','超標件數','over'])),
    fineCases:toNum(valueBy(row,['告發件數','告發','fineCases'])),
    inspectCases:toNum(valueBy(row,['通知到檢件數','通檢件數','通檢','inspectCases'])),
    fineAmount:toNum(valueBy(row,['告發金額','金額','fineAmount'])),
    completed:toBool(valueBy(row,['是否完成','完成','completed'])||true)
  };
}
function mapVehicleRow(row,index=0){
  const sourceDate=normalizeDate(valueBy(row,['日期','date','案件日期','量測日期']));
  const datetime=normalizeDateTime(valueBy(row,['量測時間','日期時間','datetime','時間','稽查時段']),sourceDate);
  const date=/^\d{4}-\d{2}-\d{2}/.test(datetime)?datetime.slice(0,10):sourceDate;
  const location=String(valueBy(row,['點位地址','點位','地址','location','量測位置地點'])||'').trim();
  const district=String(valueBy(row,['行政區','district','區域','便於複製(行政區)'])||extractDistrict(location)).trim();
  const db=toNum(valueBy(row,['量測值','分貝值','db','背景修正後分貝']));
  const standard=toNum(valueBy(row,['標準值','標準','standard','管制標準']));
  const exceedInput=valueBy(row,['超標值','超標分貝','exceed']);
  const caseType=String(valueBy(row,['案件類型','caseType','類型','_caseType'])||'').trim();
  return {
    caseType,
    plate:String(valueBy(row,['車牌','plate','車牌號碼','車號'])||'').trim().toUpperCase(),
    vehicleType:String(valueBy(row,['車種','vehicleType','車輛種類'])||'').trim(),
    date,
    datetime,
    location,
    district,
    road:String(valueBy(row,['道路','road','路段','便於複製(路段)'])||stripDistrict(location,district)).trim(),
    db, standard,
    exceed:exceedInput===''?Math.round((db-standard)*10)/10:toNum(exceedInput),
    amount:caseType==='通檢'?0:toNum(valueBy(row,['金額','罰鍰','amount'])),
    caseNo:String(valueBy(row,['案件編號','案號','caseNo','告發單號','稽查編號'])||'').trim(),
    officialRepeat:String(valueBy(row,['官方註記','是否累犯','officialRepeat'])||'').trim(),
    sourceNote:String(valueBy(row,['來源備註','備註','承辦複核','sourceNote'])||'').trim()
  };
}
function rawToRow(r){ return [r.seq??'',r.machine||'',r.caseType||'',r.period||'',r.date||'',r.month??'',r.district||'',r.location||'',r.vehicles??0,r.over??0,r.fineCases??0,r.inspectCases??0,r.fineAmount??0,r.completed?'是':'否']; }
function vehicleToRow(v){ return [v.caseType||'',v.plate||'',v.vehicleType||'',v.date||'',v.datetime||'',v.district||'',v.location||'',v.road||'',v.db??0,v.standard??0,v.exceed??0,v.amount??0,v.caseNo||'',v.officialRepeat||'',v.sourceNote||'']; }
function rowsToObjects(values){
  if(!values || values.length<2)return [];
  const headers=values[0];
  return values.slice(1).filter(r => r && r.some(c=>String(c??'').trim()!=='')).map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
}
async function readSheet(name, range='A:Z'){
  await ensureNamedSheetExists(name);
  const json=await googleSheetsRequest('GET',`https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${encRange(`'${name}'!${range}`)}`);
  return json.values || [];
}
async function writeSheet(name, columns, rows){
  await ensureNamedSheetExists(name);
  await googleSheetsRequest('POST',`https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${encRange(`'${name}'!A:Z`)}:clear`,{});
  await googleSheetsRequest('PUT',`https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${encRange(`'${name}'!A1`)}?valueInputOption=USER_ENTERED`,{values:[columns,...rows]});
}
async function pullFromGoogleSheet(){
  if(!GOOGLE_SHEET_ID || !GOOGLE_SERVICE_ACCOUNT_JSON) throw new Error('Google Sheet 尚未設定。');
  const rawValues=await readSheet(SHEET_RAW,'A:N');
  const vehicleValues=await readSheet(SHEET_VEHICLES,'A:O');
  const raw=rowsToObjects(rawValues).map(mapRawRow).filter(r=>r.date && r.location);
  const vehicles=rowsToObjects(vehicleValues).map(mapVehicleRow).filter(v=>v.caseType && v.plate && v.date);
  const data={raw,vehicles,updatedAt:new Date().toISOString(),source:'google-sheet'};
  writeData(data);
  return data;
}
async function pushToGoogleSheet(data){
  if(!GOOGLE_SHEET_ID || !GOOGLE_SERVICE_ACCOUNT_JSON) throw new Error('Google Sheet 尚未設定。');
  const raw=(data.raw||[]).map(mapRawRow);
  const vehicles=(data.vehicles||[]).map(mapVehicleRow);
  await writeSheet(SHEET_RAW, RAW_COLUMNS, raw.map(rawToRow));
  await writeSheet(SHEET_VEHICLES, VEHICLE_COLUMNS, vehicles.map(vehicleToRow));
  const out={raw,vehicles,updatedAt:new Date().toISOString(),source:'platform'};
  writeData(out);
  return {ok:true, raw:raw.length, vehicles:vehicles.length, updatedAt:out.updatedAt};
}
async function saveDashboardData(data, source='platform'){
  const raw=(data.raw||[]).map(mapRawRow);
  const vehicles=(data.vehicles||[]).map(mapVehicleRow);
  if(GOOGLE_AUTO_SYNC && googleSheetConfigured()) return await pushToGoogleSheet({raw,vehicles});
  const out={raw,vehicles,updatedAt:new Date().toISOString(),source};
  writeData(out);
  return {ok:true, raw:raw.length, vehicles:vehicles.length, updatedAt:out.updatedAt, storage:'zeabur-local'};
}
async function getDashboardData(){
  if(GOOGLE_AUTO_SYNC && googleSheetConfigured()){
    try{return await pullFromGoogleSheet();}catch(err){console.warn('[SHEET_PULL_FAILED]',err.message)}
  }
  return readData();
}
function pct(n,d){ return d ? Math.round((Number(n||0)/Number(d||0))*10000)/100 : 0; }
function vehicleCaseType(v){ return String(v?.caseType || v?.['案件類型'] || v?.type || '').trim(); }
function isFineVehicle(v){ return vehicleCaseType(v).includes('告發'); }
function isInspectVehicle(v){ return vehicleCaseType(v).includes('通檢'); }
function vehicleStats(vehicles){
  const fine=(vehicles||[]).filter(isFineVehicle).length;
  const inspect=(vehicles||[]).filter(isInspectVehicle).length;
  return {fine,inspect,total:fine+inspect};
}
function summary(data){
  const raw=data.raw||[], vehicles=data.vehicles||[];
  const completed=raw.filter(r=>r.completed!==false);
  const vehicleCounts=vehicleStats(vehicles);
  const fineVehicles=vehicleCounts.fine;
  const inspectVehicles=vehicleCounts.inspect;
  const sum=(arr,k)=>arr.reduce((s,r)=>s+(Number(r[k])||0),0);
  const vehicleDetected=sum(raw,'vehicles');
  const over=sum(raw,'over');
  // 成案件數以車輛案件明細為準，避免執行成效表彙總欄位重複累加。
  const fineCases=fineVehicles;
  const inspectCases=inspectVehicles;
  const caseTotal=vehicleCounts.total;
  const byDistrict={};
  for(const r of raw){
    const d=r.district||'未填';
    byDistrict[d]=byDistrict[d]||{sessions:0,vehicles:0,over:0,fine:0,inspect:0,caseTotal:0,kpi:0,overRate:0,fineRate:0,inspectRate:0,caseRate:0};
    byDistrict[d].sessions++;
    byDistrict[d].vehicles+=Number(r.vehicles)||0;
    byDistrict[d].over+=Number(r.over)||0;
  }
  for(const v of vehicles){
    const d=v.district||'未填';
    byDistrict[d]=byDistrict[d]||{sessions:0,vehicles:0,over:0,fine:0,inspect:0,caseTotal:0,kpi:0,overRate:0,fineRate:0,inspectRate:0,caseRate:0};
    if(isFineVehicle(v))byDistrict[d].fine++;
    else if(isInspectVehicle(v))byDistrict[d].inspect++;
  }
  Object.values(byDistrict).forEach(x=>{
    x.caseTotal=x.fine+x.inspect;
    x.kpi=x.sessions?Math.round((x.caseTotal/x.sessions)*100)/100:0;
    x.overRate=pct(x.over,x.vehicles);
    x.fineRate=pct(x.fine,x.vehicles);
    x.inspectRate=pct(x.inspect,x.vehicles);
    x.caseRate=pct(x.caseTotal,x.vehicles);
  });
  return {
    sessions: raw.length,
    completed: completed.length,
    vehicleDetected,
    over,
    fineCases,
    inspectCases,
    caseTotal,
    fineAmount: sum(raw,'fineAmount'),
    vehicleRows: vehicles.length,
    fineVehicles,
    inspectVehicles,
    kpi: completed.length?Math.round((caseTotal/completed.length)*100)/100:0,
    overRate:pct(over,vehicleDetected),
    fineRate:pct(fineCases,vehicleDetected),
    inspectRate:pct(inspectCases,vehicleDetected),
    caseRate:pct(caseTotal,vehicleDetected),
    byDistrict,
    updatedAt: data.updatedAt || ''
  };
}
function summaryMessage(data){
  const s=summary(data);
  const p=v=>`${Number(v||0).toFixed(2)}%`;
  return [
    '📊 聲音照相平台執行進度',
    `場次：${s.sessions} 場（完成 ${s.completed} 場）`,
    `辨識車流：${s.vehicleDetected.toLocaleString('zh-TW')}`,
    `超標數：${s.over.toLocaleString('zh-TW')}｜超標率：${p(s.overRate)}`,
    `告發件數：${s.fineCases.toLocaleString('zh-TW')}｜告發率：${p(s.fineRate)}`,
    `通知到檢：${s.inspectCases.toLocaleString('zh-TW')}｜通檢率：${p(s.inspectRate)}`,
    `成案件數：${s.caseTotal.toLocaleString('zh-TW')}｜成案率：${p(s.caseRate)}`,
    `場次 KPI：${s.kpi} 件/場`,
    `車輛明細：${s.vehicleRows.toLocaleString('zh-TW')} 筆（告發 ${s.fineVehicles}／通檢 ${s.inspectVehicles}）`,
    `最後更新：${s.updatedAt ? new Date(s.updatedAt).toLocaleString('zh-TW') : '-'}`
  ].join('\n');
}
function parseWorkbook(buffer){
  const wb=XLSX.read(buffer,{type:'buffer',cellDates:true});
  const out={raw:[],vehicles:[]};
  for(const name of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:'',raw:false});
    if(!rows.length)continue;
    const norm=normalizeHeader(name);
    const headers=Object.keys(rows[0]||{}).map(normalizeHeader).join('|');
    if(norm.includes('車輛')||norm.includes('案件')||headers.includes('車牌')) out.vehicles.push(...rows.map(mapVehicleRow));
    else out.raw.push(...rows.map(mapRawRow));
  }
  return out;
}
function workbookBuffer(data){
  const wb=XLSX.utils.book_new();
  const raw=(data.raw||[]).map(mapRawRow);
  const vehicles=(data.vehicles||[]).map(mapVehicleRow);
  const rawRows=raw.map(r=>Object.fromEntries(RAW_COLUMNS.map((h,i)=>[h, rawToRow(r)[i]])));
  const vehicleRows=vehicles.map(v=>Object.fromEntries(VEHICLE_COLUMNS.map((h,i)=>[h, vehicleToRow(v)[i]])));
  const s=summary({raw,vehicles,updatedAt:data.updatedAt});
  const p=v=>`${Number(v||0).toFixed(2)}%`;
  const summaryRows=[
    ['項目','數值','說明'],
    ['執行場次',s.sessions,'總匯入場次'],
    ['完成場次',s.completed,'是否完成不為否之場次'],
    ['辨識車流',s.vehicleDetected,'各場次辨識車流合計'],
    ['超標數',s.over,'各場次超標數合計'],
    ['超標率',p(s.overRate),'超標數 ÷ 辨識車流'],
    ['告發件數',s.fineCases,'各場次告發件數合計'],
    ['告發率',p(s.fineRate),'告發件數 ÷ 辨識車流'],
    ['通知到檢件數',s.inspectCases,'各場次通知到檢件數合計'],
    ['通檢率',p(s.inspectRate),'通知到檢件數 ÷ 辨識車流'],
    ['成案件數',s.caseTotal,'告發件數 + 通知到檢件數'],
    ['成案率',p(s.caseRate),'成案件數 ÷ 辨識車流'],
    ['場次 KPI',s.kpi,'成案件數 ÷ 完成場次'],
    ['告發金額',s.fineAmount,'告發金額合計'],
    ['車輛明細筆數',s.vehicleRows,'車輛案件資料筆數'],
    ['最後更新',s.updatedAt ? new Date(s.updatedAt).toLocaleString('zh-TW') : '-', '平台資料更新時間']
  ];
  const districtRows=[['行政區','場次','辨識車流','超標數','告發','通檢','成案件','KPI(件/場)','超標率','告發率','通檢率','成案率'],
    ...Object.entries(s.byDistrict).sort((a,b)=>b[1].caseTotal-a[1].caseTotal).map(([d,x])=>[d,x.sessions,x.vehicles,x.over,x.fine,x.inspect,x.caseTotal,x.kpi,p(x.overRate),p(x.fineRate),p(x.inspectRate),p(x.caseRate)])
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summaryRows),'總表KPI');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(districtRows),'行政區KPI');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rawRows,{header:RAW_COLUMNS}),SHEET_RAW);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(vehicleRows,{header:VEHICLE_COLUMNS}),SHEET_VEHICLES);
  return XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
}
async function replyLine(replyToken,text){
  if(!LINE_CHANNEL_ACCESS_TOKEN || !replyToken){
    console.warn('[LINE_REPLY_SKIPPED]', { hasToken:Boolean(LINE_CHANNEL_ACCESS_TOKEN), hasReplyToken:Boolean(replyToken) });
    return {ok:false,skipped:true};
  }
  const res = await fetch('https://api.line.me/v2/bot/message/reply', {
    method:'POST',
    headers:{Authorization:`Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,'Content-Type':'application/json'},
    body:JSON.stringify({replyToken,messages:[{type:'text',text:String(text || '').slice(0,4900)}]})
  });
  const body = await res.text();
  if(!res.ok){
    console.warn('[LINE_REPLY_FAILED]', res.status, body);
    return {ok:false,status:res.status,body};
  }
  console.log('[LINE_REPLY_OK]');
  return {ok:true};
}
function verifyLineSignature(req){
  if(!LINE_WEBHOOK_VERIFY)return true;
  const sig=req.get('x-line-signature')||'';
  const expected=crypto.createHmac('sha256',LINE_CHANNEL_SECRET).update(req.rawBody||'').digest('base64');
  try{return crypto.timingSafeEqual(Buffer.from(sig),Buffer.from(expected));}catch{return false;}
}

const ROLE_LEVEL = { viewer: 0, editor: 1, admin: 2 };
function getDashboardToken(req){
  const auth = req.get('authorization') || '';
  const bearer = auth.match(/^Bearer\s+(.+)$/i)?.[1] || '';
  return String(req.get('x-dashboard-token') || req.get('x-admin-token') || bearer || req.query.token || req.query.auth || '').trim();
}
function safeTokenEqual(a,b){
  if(!a || !b) return false;
  const ba = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  if(ba.length !== bb.length) return false;
  return crypto.timingSafeEqual(ba,bb);
}
function roleFromToken(token){
  if(token && DASHBOARD_ADMIN_TOKEN && safeTokenEqual(token,DASHBOARD_ADMIN_TOKEN)) return 'admin';
  if(token && DASHBOARD_EDITOR_TOKEN && safeTokenEqual(token,DASHBOARD_EDITOR_TOKEN)) return 'editor';
  return 'viewer';
}
function requireDashboardRole(required='editor'){
  return (req,res,next)=>{
    const token = getDashboardToken(req);
    const role = roleFromToken(token);
    const tokenConfigured = Boolean(DASHBOARD_ADMIN_TOKEN || DASHBOARD_EDITOR_TOKEN);
    if(!tokenConfigured){
      return res.status(503).json({ok:false,message:'管理權限尚未設定。請於 Zeabur 環境變數設定 DASHBOARD_ADMIN_TOKEN，必要時加設 DASHBOARD_EDITOR_TOKEN。'});
    }
    if(ROLE_LEVEL[role] >= ROLE_LEVEL[required]){
      req.dashboardRole = role;
      return next();
    }
    return res.status(401).json({ok:false,message:'權限不足或管理 Token 不正確。'});
  };
}


// ===== AI decision assistant v1.0 =====
const AI_ENABLED = String(process.env.AI_ENABLED || 'true').toLowerCase() === 'true';
const AI_PROVIDER = String(process.env.AI_PROVIDER || (process.env.OPENAI_API_KEY || process.env.AI_API_KEY ? 'openai' : 'heuristic')).toLowerCase();
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || process.env.AI_API_KEY || '';
const AI_MODEL = process.env.AI_MODEL || 'gpt-4.1-mini';
const AI_REQUIRE_ROLE = process.env.AI_REQUIRE_ROLE || 'editor';
const AI_MASK_PLATE_FOR_VIEWER = String(process.env.AI_MASK_PLATE_FOR_VIEWER || 'true').toLowerCase() !== 'false';
const AI_MAX_ROWS = Math.max(50, Math.min(3000, Number(process.env.AI_MAX_ROWS || 500)));
const AI_LOG_ENABLED = String(process.env.AI_LOG_ENABLED || 'true').toLowerCase() === 'true';

function requireAIRole(req,res,next){
  if(!AI_ENABLED) return res.status(503).json({ok:false,message:'AI 功能尚未啟用，請設定 AI_ENABLED=true。'});
  const required = ['viewer','editor','admin'].includes(AI_REQUIRE_ROLE) ? AI_REQUIRE_ROLE : 'editor';
  if(required === 'viewer') return next();
  return requireDashboardRole(required)(req,res,next);
}
function maskPlate(plate){
  const p=String(plate||'').trim().toUpperCase();
  if(p.length<=3) return p ? '***' : '';
  return `${p.slice(0,2)}***${p.slice(-2)}`;
}
function rankDistricts(s, limit=8){
  return Object.entries(s.byDistrict||{}).map(([district,x])=>({district,...x}))
    .sort((a,b)=>b.caseTotal-a.caseTotal || b.kpi-a.kpi || b.over-a.over || b.vehicles-a.vehicles)
    .slice(0,limit);
}
function rankMachines(raw, limit=8){
  return groupRows(raw||[],r=>r.machine,limit).filter(x=>x.name&&x.name!=='未填')
    .sort((a,b)=>b.caseTotal-a.caseTotal || b.kpi-a.kpi || b.vehicles-a.vehicles).slice(0,limit);
}
function buildRecommendations(data, limit=10){
  const s=summary(data);
  const districts=Object.entries(s.byDistrict||{}).map(([district,x])=>{
    const score = Math.round((x.caseTotal*8) + (x.kpi*25) + (x.overRate*1.5) + Math.min(x.vehicles/1000,30) + Math.min(x.sessions,20));
    const reasons=[];
    if(x.caseTotal>0) reasons.push(`成案件 ${x.caseTotal} 件`);
    if(x.kpi>0) reasons.push(`KPI ${x.kpi} 件/場`);
    if(x.overRate>0) reasons.push(`超標率 ${x.overRate.toFixed(2)}%`);
    if(x.vehicles>0) reasons.push(`車流 ${nfmt(x.vehicles)}`);
    const risk=x.sessions>=20 && x.kpi<0.2 ? '場次多但 KPI 偏低，建議檢討時段、架設角度或風速條件。' : (x.sessions<3 ? '資料量偏少，建議先補足樣本後再判斷。' : '可列入下期排場評估。');
    return {district,score,sessions:x.sessions,vehicles:x.vehicles,over:x.over,fine:x.fine,inspect:x.inspect,caseTotal:x.caseTotal,kpi:x.kpi,overRate:x.overRate,caseRate:x.caseRate,reasons,risk,priority:score>=80?'S':score>=55?'A':score>=30?'B':'C'};
  }).sort((a,b)=>b.score-a.score || b.caseTotal-a.caseTotal).slice(0,limit);
  return districts;
}
function buildAnomalies(data){
  const raw=(data.raw||[]).map(mapRawRow);
  const vehicles=(data.vehicles||[]).map(mapVehicleRow);
  const issues=[];
  const seqMap=new Map();
  for(const r of raw){
    const key=String(r.seq||'').trim();
    if(key){
      const prev=seqMap.get(key);
      if(prev && (prev.date!==r.date || prev.machine!==r.machine || prev.location!==r.location)) issues.push({level:'yellow',type:'場次編號重複',message:`場次 ${key} 重複，但日期、機台或點位不同，建議匯入後保留警示，不要直接覆蓋。`,ref:key});
      seqMap.set(key,r);
    }
    if(r.completed!==false && Number(r.vehicles||0)===0) issues.push({level:'red',type:'車流為 0',message:`場次 ${r.seq||'-'} 已完成但辨識車流為 0。`,ref:r.seq});
    if(Number(r.over||0)>0 && Number(r.fineCases||0)+Number(r.inspectCases||0)===0) issues.push({level:'yellow',type:'超標未成案',message:`場次 ${r.seq||'-'} 超標 ${r.over} 件，但告發/通檢為 0，建議複核案件狀態。`,ref:r.seq});
    if(r.machine && !/(OE[_-]?)?ZB0?([1-9]|10)$/i.test(String(r.machine))) issues.push({level:'yellow',type:'機台編號格式',message:`場次 ${r.seq||'-'} 機台編號 ${r.machine} 非 ZB001～ZB010 常用格式。`,ref:r.seq});
    if(r.location && r.district && !String(r.location).includes(r.district.replace('區','')) && !String(r.location).startsWith(r.district)) issues.push({level:'yellow',type:'行政區與地址待確認',message:`場次 ${r.seq||'-'} 行政區為 ${r.district}，請確認地址是否一致。`,ref:r.seq});
  }
  const plateMap=new Map();
  for(const v of vehicles){
    if(!v.plate) issues.push({level:'red',type:'車牌缺漏',message:`${v.date||'-'} ${v.district||''} 有車輛案件缺少車牌。`,ref:v.date});
    if(v.db && v.standard && Math.abs((v.db-v.standard)-(v.exceed||0))>0.6) issues.push({level:'yellow',type:'超標值不一致',message:`車牌 ${maskPlate(v.plate)} 量測值、標準值與超標值不一致，建議複核。`,ref:v.caseNo||v.plate});
    if(v.caseType && !/(告發|通檢)/.test(v.caseType)) issues.push({level:'yellow',type:'案件類型待確認',message:`車牌 ${maskPlate(v.plate)} 案件類型為「${v.caseType}」，非告發/通檢常用分類。`,ref:v.caseNo||v.plate});
    if(v.plate){
      const key=normalizePlateToken(v.plate);
      plateMap.set(key,(plateMap.get(key)||0)+1);
    }
  }
  for(const [plate,count] of plateMap.entries()) if(count>=3) issues.push({level:'yellow',type:'高頻車牌',message:`車牌 ${maskPlate(plate)} 於案件明細出現 ${count} 次，可列入複核或追蹤。`,ref:plate});
  const red=issues.filter(x=>x.level==='red').length;
  const yellow=issues.filter(x=>x.level==='yellow').length;
  return {red,yellow,green:Math.max(0,raw.length+vehicles.length-red-yellow),issues:issues.slice(0,80)};
}
function buildAIContext(data, options={}){
  const role=options.role||'viewer';
  const s=summary(data);
  const raw=(data.raw||[]).map(mapRawRow).slice(0,AI_MAX_ROWS);
  const vehicles=(data.vehicles||[]).map(mapVehicleRow).slice(0,AI_MAX_ROWS).map(v=>AI_MASK_PLATE_FOR_VIEWER&&role==='viewer'?{...v,plate:maskPlate(v.plate),caseNo:v.caseNo?'已遮罩':''}:v);
  return {summary:s,topDistricts:rankDistricts(s,10),topMachines:rankMachines(raw,10),recommendations:buildRecommendations({raw,vehicles},10),anomalies:buildAnomalies({raw,vehicles}),sample:{raw:raw.slice(0,30),vehicles:vehicles.slice(0,30)}};
}
function formatRecommendations(recs){
  if(!recs.length) return '目前資料不足，無法產生排場建議。';
  return recs.map((x,i)=>`${i+1}. ${x.district}｜${x.priority}級｜分數 ${x.score}｜場次 ${x.sessions}｜成案 ${x.caseTotal}｜KPI ${x.kpi}\n   原因：${x.reasons.join('、')||'資料量不足'}\n   風險：${x.risk}`).join('\n');
}
function heuristicReport(data, style='主管簡報版'){
  const s=summary(data);
  const top=rankDistricts(s,5);
  const recs=buildRecommendations(data,5);
  const anomalies=buildAnomalies(data);
  const topText=top.map((x,i)=>`${i+1}. ${x.district}：場次 ${x.sessions}、告發 ${x.fine}、通檢 ${x.inspect}、成案 ${x.caseTotal}、KPI ${x.kpi}`).join('\n') || '目前無行政區資料。';
  const keyFinding=[];
  if(top[0]) keyFinding.push(`${top[0].district}為目前成案件數最高行政區，成案 ${top[0].caseTotal} 件，KPI ${top[0].kpi} 件/場。`);
  const lowKpi=top.find(x=>x.sessions>=10 && x.kpi<0.2);
  if(lowKpi) keyFinding.push(`${lowKpi.district}場次已有 ${lowKpi.sessions} 場但 KPI 偏低，建議檢討架設條件與執行時段。`);
  if(anomalies.red||anomalies.yellow) keyFinding.push(`資料健檢發現紅色 ${anomalies.red} 件、黃色 ${anomalies.yellow} 件，建議匯出前先複核。`);
  if(!keyFinding.length) keyFinding.push('目前資料未見明顯異常，可持續觀察行政區與機台 KPI 變化。');
  return [
    `AI 成效摘要（${style}）`,
    '',
    '一、總體成果',
    `累計執行 ${s.sessions} 場，完成 ${s.completed} 場；辨識車流 ${nfmt(s.vehicleDetected)} 輛次，超標 ${nfmt(s.over)} 件，告發 ${nfmt(s.fineCases)} 件、通知到檢 ${nfmt(s.inspectCases)} 件，整體成案件數 ${nfmt(s.caseTotal)} 件，場次 KPI ${s.kpi} 件/場。`,
    '',
    '二、關鍵判讀',
    keyFinding.map((x,i)=>`${i+1}. ${x}`).join('\n'),
    '',
    '三、行政區重點排行',
    topText,
    '',
    '四、下期排場建議',
    formatRecommendations(recs),
    '',
    '五、風險提醒',
    anomalies.issues.slice(0,5).map((x,i)=>`${i+1}. [${x.level==='red'?'紅':'黃'}] ${x.message}`).join('\n') || '目前無重大資料異常。',
    '',
    `最後更新：${s.updatedAt ? new Date(s.updatedAt).toLocaleString('zh-TW') : '-'}`,
    '註：本摘要依平台現有資料自動產出，涉及裁罰與正式行政處分仍應以原始表單及承辦複核為準。'
  ].join('\n');
}
async function callOpenAIForText(task, context, question){
  if(!OPENAI_API_KEY || AI_PROVIDER !== 'openai') return '';
  const system = '你是新北市噪音車科技執法資料分析助理。只能根據使用者提供的 JSON 資料回答，不得編造數字。輸出繁體中文，分為摘要、關鍵數據、判讀、建議、風險提醒。涉及裁罰與正式行政處分必須提醒仍以承辦複核為準。';
  const body={model:AI_MODEL,messages:[{role:'system',content:system},{role:'user',content:JSON.stringify({task,question,context})}],temperature:0.2};
  const res=await fetch('https://api.openai.com/v1/chat/completions',{method:'POST',headers:{Authorization:`Bearer ${OPENAI_API_KEY}`,'Content-Type':'application/json'},body:JSON.stringify(body)});
  const json=await res.json().catch(()=>({}));
  if(!res.ok) throw new Error(json?.error?.message || `OpenAI API ${res.status}`);
  return json?.choices?.[0]?.message?.content || '';
}
async function generateAIText(task, data, question='', role='viewer', fallbackStyle='主管簡報版'){
  const context=buildAIContext(data,{role});
  if(AI_LOG_ENABLED) console.log('[AI_TASK]',task,{provider:AI_PROVIDER,openai:Boolean(OPENAI_API_KEY),question:String(question||'').slice(0,80)});
  try{
    const out=await callOpenAIForText(task, context, question);
    if(out) return {provider:'openai',text:out,context};
  }catch(err){
    console.warn('[AI_OPENAI_FALLBACK]',err.message);
  }
  let text='';
  if(task==='recommend') text=`AI 排場建議\n\n${formatRecommendations(context.recommendations)}\n\n註：此排序依現有資料、成案件數、KPI、超標率、車流與場次量推估，正式排場仍需現勘確認。`;
  else if(task==='validate') text=`AI 資料健檢\n\n紅色：${context.anomalies.red} 件\n黃色：${context.anomalies.yellow} 件\n可用資料：${context.anomalies.green} 筆\n\n${context.anomalies.issues.slice(0,20).map((x,i)=>`${i+1}. [${x.level==='red'?'紅':'黃'}] ${x.type}：${x.message}`).join('\n') || '目前未發現明顯異常。'}`;
  else if(task==='query'){
    const q=parseLineQuery(question||'');
    if(hasLineQuery(q)) text=lineFilteredMessage(data,q);
    else text=heuristicReport(data,'AI 問答版');
  }else text=heuristicReport(data,fallbackStyle);
  return {provider:'heuristic',text,context};
}

app.use(express.json({limit:'50mb',verify:(req,res,buf)=>{req.rawBody=buf}}));
app.use(express.urlencoded({extended:true,limit:'50mb'}));
app.use(express.static(path.join(__dirname,'public')));

app.get('/healthz', (req,res)=>res.status(200).send('ok'));
app.get('/api/health', (req,res)=>res.json({
  ok:true,
  uptime:Math.round(process.uptime()),
  storage: GOOGLE_AUTO_SYNC && googleSheetConfigured() ? 'google-sheet-realtime' : 'zeabur-local',
  googleAutoSync:GOOGLE_AUTO_SYNC,
  googleSheetConfigured:googleSheetConfigured(),
  googleSheetRawName:SHEET_RAW,
  googleSheetVehicleName:SHEET_VEHICLES,
  googleSheetUrl:GOOGLE_SHEET_URL,
  lineConfigured:Boolean(LINE_CHANNEL_ACCESS_TOKEN && LINE_CHANNEL_SECRET),
  tokenConfigured:Boolean(DASHBOARD_ADMIN_TOKEN || DASHBOARD_EDITOR_TOKEN),
  ai:{enabled:AI_ENABLED,provider:AI_PROVIDER,model:AI_MODEL,requireRole:AI_REQUIRE_ROLE,openaiConfigured:Boolean(OPENAI_API_KEY)},
  realtimeNote: GOOGLE_AUTO_SYNC && googleSheetConfigured() ? '前台、後台匯出與 LINE Bot 查詢時會即時讀取 Google Sheet。' : '目前未啟用 Google Sheet 即時同步，使用 Zeabur 本地快取。'
}));


app.get('/api/google-sheet/test', requireDashboardRole('editor'), async (req,res)=>{
  try{
    if(!GOOGLE_AUTO_SYNC) throw new Error('GOOGLE_AUTO_SYNC 目前不是 true，尚未啟用 Google Sheet 即時同步。');
    if(!googleSheetConfigured()) throw new Error('GOOGLE_SHEET_ID 或 GOOGLE_SERVICE_ACCOUNT_JSON 尚未設定。');
    const rawValues = await readSheet(SHEET_RAW,'A:N');
    const vehicleValues = await readSheet(SHEET_VEHICLES,'A:O');
    const rawRows = rowsToObjects(rawValues).map(mapRawRow).filter(r=>r.date && r.location);
    const vehicleRows = rowsToObjects(vehicleValues).map(mapVehicleRow).filter(v=>v.caseType && v.plate && v.date);
    res.json({
      ok:true,
      googleAutoSync:GOOGLE_AUTO_SYNC,
      googleSheetConfigured:googleSheetConfigured(),
      sheetId:GOOGLE_SHEET_ID,
      sheetUrl:GOOGLE_SHEET_URL,
      sheets:{
        raw:{name:SHEET_RAW, rows:rawRows.length, header:(rawValues[0]||[])},
        vehicles:{name:SHEET_VEHICLES, rows:vehicleRows.length, header:(vehicleValues[0]||[])}
      },
      summary:summary({raw:rawRows,vehicles:vehicleRows,updatedAt:new Date().toISOString()}),
      message:'Google Sheet 讀取測試成功。平台、LINE Bot 與匯出功能會以這份 Sheet 作為即時資料來源。'
    });
  }catch(err){
    res.status(400).json({ok:false,message:err.message,googleAutoSync:GOOGLE_AUTO_SYNC,googleSheetConfigured:googleSheetConfigured(),sheetIdConfigured:Boolean(GOOGLE_SHEET_ID),serviceAccountConfigured:Boolean(GOOGLE_SERVICE_ACCOUNT_JSON)});
  }
});

app.post('/api/google-sheet/refresh', requireDashboardRole('editor'), async (req,res)=>{
  try{
    const data = await pullFromGoogleSheet();
    res.json({ok:true,raw:data.raw.length,vehicles:data.vehicles.length,updatedAt:data.updatedAt,summary:summary(data),message:'已從 Google Sheet 重新讀取並更新 Zeabur 快取。'});
  }catch(err){
    res.status(400).json({ok:false,message:err.message});
  }
});

app.get(['/admin','/management'], (req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));

app.get('/api/auth/check', (req,res)=>{
  const token = getDashboardToken(req);
  const role = roleFromToken(token);
  const tokenConfigured = Boolean(DASHBOARD_ADMIN_TOKEN || DASHBOARD_EDITOR_TOKEN);
  if(!tokenConfigured) return res.status(503).json({ok:false,role:'viewer',message:'管理權限尚未設定。'});
  if(token && role === 'viewer') return res.status(401).json({ok:false,role:'viewer',message:'管理 Token 不正確。'});
  res.json({
    ok:true,
    role,
    permissions:{
      view:true,
      importData: ROLE_LEVEL[role] >= ROLE_LEVEL.editor,
      exportData: ROLE_LEVEL[role] >= ROLE_LEVEL.editor,
      syncData: ROLE_LEVEL[role] >= ROLE_LEVEL.editor,
      restoreBuiltIn: ROLE_LEVEL[role] >= ROLE_LEVEL.admin
    }
  });
});

app.get('/api/noise-dashboard-data', async (req,res)=>{
  try{const data=await getDashboardData();res.json({...data,summary:summary(data)});}catch(err){res.status(500).json({ok:false,message:err.message})}
});
app.post('/api/noise-dashboard-data', requireDashboardRole('editor'), async (req,res)=>{
  try{
    const payload={raw:Array.isArray(req.body.raw)?req.body.raw:[],vehicles:Array.isArray(req.body.vehicles)?req.body.vehicles:[],updatedAt:new Date().toISOString(),source:req.body.source||'api'};
    const result=await saveDashboardData(payload,'api');
    res.json({ok:true,result,summary:summary(readData())});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});
app.post('/api/sync/pull', requireDashboardRole('editor'), async (req,res)=>{try{const data=await pullFromGoogleSheet();res.json({ok:true,raw:data.raw.length,vehicles:data.vehicles.length,summary:summary(data)})}catch(err){res.status(400).json({ok:false,message:err.message})}});
app.post('/api/sync/push', requireDashboardRole('admin'), async (req,res)=>{try{const result=await pushToGoogleSheet(readData());res.json({ok:true,result})}catch(err){res.status(400).json({ok:false,message:err.message})}});
app.post('/api/import/excel', requireDashboardRole('editor'), upload.single('file'), async (req,res)=>{
  try{if(!req.file)throw new Error('請上傳 Excel 檔案。');const parsed=parseWorkbook(req.file.buffer);const result=await saveDashboardData(parsed,'excel-import');res.json({ok:true,result,summary:summary(readData())});}catch(err){res.status(400).json({ok:false,message:err.message})}
});
app.post('/api/restore-built-in', requireDashboardRole('admin'), async (req,res)=>{
  try{
    const payload={raw:Array.isArray(req.body.raw)?req.body.raw:[],vehicles:Array.isArray(req.body.vehicles)?req.body.vehicles:[],updatedAt:new Date().toISOString(),source:'restore-built-in'};
    const result=await saveDashboardData(payload,'restore-built-in');
    res.json({ok:true,result,summary:summary(readData())});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});
app.get('/api/export/excel', requireDashboardRole('editor'), async (req,res)=>{
  try{
    // 匯出前先讀取 Google Sheet，確保匯出內容與目前匯入／同步後資料一致。
    const data = await getDashboardData();
    const normalized = {
      raw: (data.raw || []).map(mapRawRow),
      vehicles: (data.vehicles || []).map(mapVehicleRow),
      updatedAt: new Date().toISOString(),
      source: 'export-synced'
    };
    writeData(normalized);
    const buf = workbookBuffer(normalized);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename*=UTF-8''${encodeURIComponent('聲音照相資料同步匯出.xlsx')}`);
    res.send(buf);
  }catch(err){res.status(500).send(err.message)}
});
app.get('/api/export/summary-excel', requireDashboardRole('editor'), async (req,res)=>{
  try{
    const data = await getDashboardData();
    const normalized = {
      raw: (data.raw || []).map(mapRawRow),
      vehicles: (data.vehicles || []).map(mapVehicleRow),
      updatedAt: new Date().toISOString(),
      source: 'export-total-summary'
    };
    writeData(normalized);
    const buf = workbookBuffer(normalized);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename*=UTF-8''${encodeURIComponent('聲音照相執行總表KPI.xlsx')}`);
    res.send(buf);
  }catch(err){res.status(500).send(err.message)}
});
app.get('/api/summary', async (req,res)=>{try{const data=await getDashboardData();res.json({ok:true,summary:summary(data)})}catch(err){res.status(500).json({ok:false,message:err.message})}});

app.get('/api/ai/status', (req,res)=>res.json({
  ok:true,
  enabled:AI_ENABLED,
  provider:AI_PROVIDER,
  model:AI_MODEL,
  openaiConfigured:Boolean(OPENAI_API_KEY),
  requireRole:AI_REQUIRE_ROLE,
  maxRows:AI_MAX_ROWS,
  features:['query','report','validate-import','recommend-locations','anomaly-check']
}));

app.post('/api/ai/query', requireAIRole, async (req,res)=>{
  try{
    const data=await getDashboardData();
    const role=req.dashboardRole || roleFromToken(getDashboardToken(req));
    const question=String(req.body?.question || req.body?.q || '').trim();
    if(!question) throw new Error('請提供 question。');
    const result=await generateAIText('query',data,question,role,'AI 問答版');
    res.json({ok:true,provider:result.provider,answer:result.text,summary:result.context.summary});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});

app.post('/api/ai/report', requireAIRole, async (req,res)=>{
  try{
    const data=await getDashboardData();
    const role=req.dashboardRole || roleFromToken(getDashboardToken(req));
    const style=String(req.body?.style || process.env.AI_REPORT_DEFAULT_STYLE || '主管簡報版');
    const result=await generateAIText('report',data,`請產出${style}`,role,style);
    res.json({ok:true,provider:result.provider,style,report:result.text,summary:result.context.summary});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});

app.post('/api/ai/validate-import', requireAIRole, upload.single('file'), async (req,res)=>{
  try{
    let data;
    if(req.file) data=parseWorkbook(req.file.buffer);
    else if(req.body && (Array.isArray(req.body.raw)||Array.isArray(req.body.vehicles))) data={raw:req.body.raw||[],vehicles:req.body.vehicles||[]};
    else data=await getDashboardData();
    const role=req.dashboardRole || roleFromToken(getDashboardToken(req));
    const context=buildAIContext(data,{role});
    const result=await generateAIText('validate',data,'請檢查匯入資料異常',role,'資料健檢版');
    res.json({ok:true,provider:result.provider,validation:context.anomalies,message:result.text});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});

app.post('/api/ai/recommend-locations', requireAIRole, async (req,res)=>{
  try{
    const data=await getDashboardData();
    const role=req.dashboardRole || roleFromToken(getDashboardToken(req));
    const context=buildAIContext(data,{role});
    const result=await generateAIText('recommend',data,'請產生下期排場與百大點位排序建議',role,'排場建議版');
    res.json({ok:true,provider:result.provider,recommendations:context.recommendations,message:result.text});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});

app.post('/api/ai/anomaly-check', requireAIRole, async (req,res)=>{
  try{
    const data=await getDashboardData();
    const role=req.dashboardRole || roleFromToken(getDashboardToken(req));
    const context=buildAIContext(data,{role});
    res.json({ok:true,anomalies:context.anomalies,summary:context.summary});
  }catch(err){res.status(400).json({ok:false,message:err.message})}
});

const LINE_COMMANDS = [
  '最新進度',
  '執行進度',
  '2月份執行成效',
  '6月進度',
  '淡水區執行成效',
  '淡水區成效',
  '日期2026-06-12',
  '機號OE_ZB004',
  '車牌ABC-1234',
  '淡水區 6月 告發率',
  '查詢說明',
  'AI 幫我看本週重點',
  'AI 建議下週排場',
  'AI 檢查今天資料有沒有異常'
];

app.get('/line-webhook',(req,res)=>res.json({
  ok:true,
  message:'LINE webhook endpoint ready',
  lineConfigured:Boolean(LINE_CHANNEL_ACCESS_TOKEN && LINE_CHANNEL_SECRET),
  tokenConfigured:Boolean(LINE_CHANNEL_ACCESS_TOKEN),
  secretConfigured:Boolean(LINE_CHANNEL_SECRET),
  sheetConfigured:Boolean(GOOGLE_SHEET_ID && GOOGLE_SERVICE_ACCOUNT_JSON),
  verifySignature:LINE_WEBHOOK_VERIFY,
  mode:'google_sheet_realtime_line_detailed_kpi_v8',
  commands:LINE_COMMANDS
}));

app.get('/api/line-debug',(req,res)=>res.json({
  ok:true,
  tokenConfigured:Boolean(LINE_CHANNEL_ACCESS_TOKEN),
  secretConfigured:Boolean(LINE_CHANNEL_SECRET),
  verifySignature:LINE_WEBHOOK_VERIFY,
  sheetConfigured:Boolean(GOOGLE_SHEET_ID && GOOGLE_SERVICE_ACCOUNT_JSON),
  commands:LINE_COMMANDS,
  querySupport:['月份執行成效','日期執行成效','行政區執行成效','機號/機台','車牌號碼','KPI 指標','佐證明細','排名摘要'],
  hint:'LINE Bot 沒回應時，先確認 LINE Developers Webhook URL 是否為 /line-webhook，並查看 Zeabur Logs 是否出現 [LINE_WEBHOOK_RECEIVED]。'
}));

function isSummaryCommand(clean){
  return /^(最新進度|執行進度|成果總數|最新數據|總數|狀態|更新總數|立即回報|summary)$/i.test(String(clean || '').trim());
}
function isHelpCommand(clean){
  return /^(查詢說明|說明|help|指令|怎麼查)$/i.test(String(clean || '').trim());
}
function lineHelpMessage(){
  return [
    '可用 LINE 即時查詢平台資料：',
    '',
    '一、總表',
    '・執行進度',
    '・最新進度',
    '',
    '二、條件查詢',
    '・2月份執行成效',
    '・6月進度',
    '・2026-06-12 成效',
    '・淡水區執行成效',
    '・淡水區 成效',
    '・機號 OE_ZB004',
    '・車牌 ABC-1234',
    '・淡水區 6月 告發率',
    '',
    '三、回覆內容',
    '執行場次、完成場次、車流、超標、告發、通檢、成案、告發率、通檢率、KPI、排行與佐證明細。',
    '',
    '系統會即時讀取 Google Sheet／平台資料，並列出摘要、排行與佐證資料。'
  ].join('\n');
}
function canonicalDistrict(text){
  const m=String(text||'').match(/(萬里|金山|板橋|汐止|深坑|石碇|瑞芳|平溪|雙溪|貢寮|新店|坪林|烏來|永和|中和|土城|三峽|樹林|鶯歌|三重|新莊|泰山|林口|蘆洲|五股|八里|淡水|三芝|石門)區?/);
  return m ? `${m[1]}區` : '';
}
function normalizePlateToken(v){
  return String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
}
function extractPlateQuery(text){
  const t=String(text||'').toUpperCase();
  const explicit=t.match(/(?:車牌|車號|牌照)\s*[:：]?\s*([A-Z0-9]{2,4}[-\s]?[A-Z0-9]{2,5})/i);
  if(explicit) return normalizePlateToken(explicit[1]);
  const general=t.match(/\b([A-Z]{1,3}[-\s]?\d{3,5}|\d{3,5}[-\s]?[A-Z]{1,3}|[A-Z0-9]{2,4}[-\s][A-Z0-9]{2,5})\b/i);
  return general ? normalizePlateToken(general[1]) : '';
}
function extractMachineQuery(text){
  const t=String(text||'');
  const explicit=t.match(/(?:機號|機台|機台編號|設備|設備編號)\s*[:：]?\s*([A-Za-z0-9_\-]+)/i);
  if(explicit) return String(explicit[1]).trim().toUpperCase();
  const oe=t.match(/\b(OE[_\-]?[A-Za-z0-9]+)\b/i);
  return oe ? String(oe[1]).replace('-', '_').toUpperCase() : '';
}
function extractMonthQuery(text){
  const t=String(text||'');
  let m=t.match(/(?:^|[^\d])(?:20\d{2}|1\d{2})[年\-.\/]\s*(\d{1,2})\s*月?/);
  if(m) return Number(m[1]);
  m=t.match(/(?:^|[^\d])(\d{1,2})\s*月/);
  if(m) return Number(m[1]);
  m=t.match(/月份\s*[:：]?\s*(\d{1,2})/);
  if(m) return Number(m[1]);
  return 0;
}
function extractDateQuery(text){
  const t=String(text||'').trim();
  const nowYear=new Date().getFullYear();
  let m=t.match(/(20\d{2})[\-.\/]\s*(\d{1,2})[\-.\/]\s*(\d{1,2})/);
  if(m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  m=t.match(/(1\d{2})[年\-.\/]\s*(\d{1,2})[月\-.\/]\s*(\d{1,2})/);
  if(m) return `${Number(m[1])+1911}-${pad2(m[2])}-${pad2(m[3])}`;
  m=t.match(/(?:日期|日|date)\s*[:：]?\s*(\d{1,2})[\/.\-](\d{1,2})/i);
  if(m) return `${nowYear}-${pad2(m[1])}-${pad2(m[2])}`;
  if(/(今天|今日)/.test(t)){
    const d=new Date(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
  }
  return '';
}
function parseLineQuery(text){
  const clean=String(text||'').trim();
  return {
    text:clean,
    district:canonicalDistrict(clean),
    month:extractMonthQuery(clean),
    date:extractDateQuery(clean),
    machine:extractMachineQuery(clean),
    plate:extractPlateQuery(clean),
    wantsRate:/(率|kpi|KPI|成效|效率|統計|進度)/i.test(clean)
  };
}
function hasLineQuery(q){
  return Boolean(q.district || q.month || q.date || q.machine || q.plate || q.wantsRate);
}
function recordMonth(date, fallback){
  if(fallback) return Number(fallback)||0;
  const ds=String(date||'').trim();
  let m=ds.match(/^\d{4}-(\d{1,2})-/);
  if(m) return Number(m[1]);
  m=ds.match(/^\d{4}\/(\d{1,2})\//);
  if(m) return Number(m[1]);
  m=ds.match(/^\d{3}\.(\d{1,2})\./);
  if(m) return Number(m[1]);
  return 0;
}
function recordDateKey(date){
  const ds=String(date||'').trim();
  let m=ds.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if(m) return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  m=ds.match(/^(\d{3})\.(\d{1,2})\.(\d{1,2})/);
  if(m) return `${m[1]}.${String(m[2]).padStart(2,'0')}.${String(m[3]).padStart(2,'0')}`;
  return ds.slice(0,10);
}
function filterForLineQuery(data,q){
  let raw=(data.raw||[]).map(mapRawRow);
  let vehicles=(data.vehicles||[]).map(mapVehicleRow);
  if(q.district){ raw=raw.filter(r=>String(r.district||'')===q.district); vehicles=vehicles.filter(v=>String(v.district||'')===q.district); }
  if(q.month){ raw=raw.filter(r=>recordMonth(r.date,r.month)===q.month); vehicles=vehicles.filter(v=>recordMonth(v.date,0)===q.month); }
  if(q.date){ raw=raw.filter(r=>recordDateKey(r.date)===q.date); vehicles=vehicles.filter(v=>recordDateKey(v.date||v.datetime)===q.date); }
  if(q.machine){ raw=raw.filter(r=>String(r.machine||'').toUpperCase().replace('-', '_').includes(q.machine)); }
  if(q.plate){ vehicles=vehicles.filter(v=>normalizePlateToken(v.plate).includes(q.plate) || q.plate.includes(normalizePlateToken(v.plate))); raw = q.machine ? raw : raw; }
  return {raw,vehicles,updatedAt:data.updatedAt||''};
}
function lineQueryTitle(q){
  const parts=[];
  if(q.district) parts.push(q.district);
  if(q.month) parts.push(`${q.month}月`);
  if(q.date) parts.push(q.date);
  if(q.machine) parts.push(`機號 ${q.machine}`);
  if(q.plate) parts.push(`車牌 ${q.plate}`);
  return parts.length ? parts.join('｜') : '全資料';
}
function compactLocation(s){
  const t=String(s||'').replace(/^新北市/,'').trim();
  return t.length>24 ? `${t.slice(0,24)}…` : t;
}
function nfmt(v){ return Number(v||0).toLocaleString('zh-TW'); }
function rateText(v){ return `${Number(v||0).toFixed(2)}%`; }
function rangeText(rows){
  const dates=(rows||[]).map(r=>String(r.date||'').slice(0,10)).filter(Boolean).sort();
  if(!dates.length) return '-';
  return dates[0]===dates[dates.length-1] ? dates[0] : `${dates[0]} 至 ${dates[dates.length-1]}`;
}
function groupRows(rows,keyFn){
  const m=new Map();
  for(const r of rows||[]){
    const k=keyFn(r)||'未填';
    const x=m.get(k)||{name:k,sessions:0,vehicles:0,over:0,fine:0,inspect:0,caseTotal:0,kpi:0,caseRate:0};
    x.sessions += 1;
    x.vehicles += Number(r.vehicles)||0;
    x.over += Number(r.over)||0;
    x.fine += Number(r.fineCases)||0;
    x.inspect += Number(r.inspectCases)||0;
    m.set(k,x);
  }
  for(const x of m.values()){
    x.caseTotal=x.fine+x.inspect;
    x.kpi=x.sessions?Math.round((x.caseTotal/x.sessions)*100)/100:0;
    x.caseRate=pct(x.caseTotal,x.vehicles);
  }
  return [...m.values()];
}
function topRows(rows,keyFn,limit=5){
  return groupRows(rows,keyFn).sort((a,b)=>b.caseTotal-a.caseTotal || b.sessions-a.sessions || b.vehicles-a.vehicles).slice(0,limit);
}
function topVehicleTypes(vehicles,limit=4){
  const m=new Map();
  for(const v of vehicles||[]){
    const k=vehicleCaseType(v)||'未填';
    m.set(k,(m.get(k)||0)+1);
  }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]).slice(0,limit);
}
function lineFilteredMessage(allData,q){
  const filtered=filterForLineQuery(allData,q);
  const s=summary(filtered);
  const title=lineQueryTitle(q);
  const rawRows=filtered.raw||[];
  const vehicleRows=filtered.vehicles||[];
  const rawProof=rawRows.slice().sort((a,b)=>String(b.date||'').localeCompare(String(a.date||''))).slice(0,5);
  const vehicleProof=vehicleRows.slice().sort((a,b)=>String(b.datetime||b.date||'').localeCompare(String(a.datetime||a.date||''))).slice(0,5);
  const lines=[
    `🔎 查詢條件：${title}`,
    `資料期間：${rangeText(rawRows)}`,
    '',
    '📊 執行成效摘要',
    `執行場次：${nfmt(s.sessions)} 場`,
    `完成場次：${nfmt(s.completed)} 場`,
    `辨識車流：${nfmt(s.vehicleDetected)}`,
    `超標數：${nfmt(s.over)}（超標率 ${rateText(s.overRate)}）`,
    `告發件數：${nfmt(s.fineCases)}（告發率 ${rateText(s.fineRate)}）`,
    `通檢件數：${nfmt(s.inspectCases)}（通檢率 ${rateText(s.inspectRate)}）`,
    `成案件數：${nfmt(s.caseTotal)}（成案率 ${rateText(s.caseRate)}）`,
    `場次 KPI：${s.kpi} 件/場`,
    `告發金額：${nfmt(s.fineAmount)} 元`,
    `車輛明細：${nfmt(s.vehicleRows)} 筆（告發 ${nfmt(s.fineVehicles)}／通檢 ${nfmt(s.inspectVehicles)}）`,
    ''
  ];

  if(!q.district){
    const topDistricts=Object.entries(s.byDistrict||{}).map(([name,x])=>({name,...x})).sort((a,b)=>b.caseTotal-a.caseTotal || b.sessions-a.sessions || b.vehicles-a.vehicles).slice(0,5);
    if(topDistricts.length){
      lines.push('🏙 行政區排行（依成案件數）');
      topDistricts.forEach((x,i)=>lines.push(`${i+1}. ${x.name}｜場次 ${x.sessions}｜告發 ${x.fine}｜通檢 ${x.inspect}｜成案 ${x.caseTotal}｜KPI ${x.kpi}`));
      lines.push('');
    }
  }

  if(!q.machine){
    const topMachines=topRows(rawRows,r=>r.machine,5).filter(x=>x.name && x.name!=='未填');
    if(topMachines.length){
      lines.push('📷 機台排行（依成案件數）');
      topMachines.forEach((x,i)=>lines.push(`${i+1}. ${x.name}｜場次 ${x.sessions}｜告發 ${x.fine}｜通檢 ${x.inspect}｜成案 ${x.caseTotal}`));
      lines.push('');
    }
  }

  const vehicleTypes=topVehicleTypes(vehicleRows,4);
  if(vehicleTypes.length){
    lines.push('🚗 車輛案件類型');
    vehicleTypes.forEach(([k,v])=>lines.push(`・${k}：${nfmt(v)} 筆`));
    lines.push('');
  }

  if(q.plate){
    lines.push('🚗 車牌佐證');
    if(vehicleProof.length){
      vehicleProof.forEach((v,i)=>{
        lines.push(`${i+1}. ${v.plate || '-'}｜${v.caseType || '-'}｜${v.date || '-'} ${String(v.datetime||'').replace(v.date||'').trim()}`);
        lines.push(`   ${v.district || '-'}｜${compactLocation(v.location || v.road)}`);
        lines.push(`   量測 ${v.db || '-'} / 標準 ${v.standard || '-'} / 超標 ${v.exceed || '-'}`);
        if(v.caseNo) lines.push(`   案件編號：${v.caseNo}`);
      });
    }else{
      lines.push('查無符合車牌資料。');
    }
  }else{
    lines.push('📌 場次佐證（最近 5 筆）');
    if(rawProof.length){
      rawProof.forEach((r,i)=>{
        lines.push(`${i+1}. S${r.seq || '-'}｜${r.date || '-'}｜${r.machine || '-'}`);
        lines.push(`   ${r.district || '-'}｜${compactLocation(r.location)}`);
        lines.push(`   車流 ${nfmt(r.vehicles)}｜超標 ${nfmt(r.over)}｜告發 ${nfmt(r.fineCases)}｜通檢 ${nfmt(r.inspectCases)}`);
      });
    }else{
      lines.push('查無符合場次資料。');
    }
    if(vehicleProof.length){
      lines.push('');
      lines.push('🚗 車輛明細佐證（最近 5 筆）');
      vehicleProof.forEach((v,i)=>{
        lines.push(`${i+1}. ${v.plate || '-'}｜${v.caseType || '-'}｜${v.date || '-'}｜${v.district || '-'}`);
        lines.push(`   ${compactLocation(v.location || v.road)}｜量測 ${v.db || '-'} dB｜超標 ${v.exceed || '-'}`);
      });
    }
  }
  lines.push('');
  lines.push(`最後更新：${s.updatedAt ? new Date(s.updatedAt).toLocaleString('zh-TW') : '-'}`);
  lines.push('註：若需完整明細，請至後台一鍵產出 Excel 總表。');
  return lines.join('\n');
}


app.post('/line-webhook',async(req,res)=>{
  const events = Array.isArray(req.body?.events) ? req.body.events : [];
  console.log('[LINE_WEBHOOK_RECEIVED]', {
    events: events.length,
    verifySignature: LINE_WEBHOOK_VERIFY,
    hasToken: Boolean(LINE_CHANNEL_ACCESS_TOKEN),
    hasSecret: Boolean(LINE_CHANNEL_SECRET),
    types: events.map(e => e.type + ':' + (e.message?.type || ''))
  });

  // 先回 200，避免 LINE Developers 判定 Webhook 失敗。
  res.status(200).json({ok:true});

  try{
    if(LINE_WEBHOOK_VERIFY && !verifyLineSignature(req)){
      console.warn('[LINE_SIGNATURE_WARNING]', '簽章驗證未通過；內部控管版仍繼續處理，避免 Bot 無回應。請確認 LINE_CHANNEL_SECRET 是否正確。');
    }

    for(const event of events){
      if(event.type !== 'message') continue;
      if(event.message?.type !== 'text'){
        await replyLine(event.replyToken,'目前僅支援文字查詢。請輸入「最新進度」、「執行進度」、「成果總數」或「最新數據」。');
        continue;
      }

      const clean = String(event.message.text || '').trim();
      try{
        if(isSummaryCommand(clean)){
          const data = await getDashboardData();
          await replyLine(event.replyToken, summaryMessage(data));
        }else if(isHelpCommand(clean)){
          await replyLine(event.replyToken, lineHelpMessage());
        }else if(/^AI[\s：:]/i.test(clean) || /^(幫我看本週重點|建議下週排場|檢查今天資料有沒有異常|幫我寫主管摘要)$/i.test(clean)){
          const data = await getDashboardData();
          const aiQuestion = clean.replace(/^AI[\s：:]*/i,'').trim() || '請產生主管摘要';
          const task = /排場|點位|建議/.test(aiQuestion) ? 'recommend' : (/異常|檢查|健檢/.test(aiQuestion) ? 'validate' : 'query');
          const result = await generateAIText(task,data,aiQuestion,'editor','LINE AI 版');
          await replyLine(event.replyToken, result.text);
        }else{
          const q=parseLineQuery(clean);
          if(hasLineQuery(q)){
            const data = await getDashboardData();
            await replyLine(event.replyToken, lineFilteredMessage(data,q));
          }else{
            await replyLine(event.replyToken, lineHelpMessage());
          }
        }
      }catch(err){
        console.error('[LINE_COMMAND_ERROR]', err);
        await replyLine(event.replyToken, [
          '❌ 查詢失敗',
          `原因：${err.message || err}`,
          '',
          '請確認 Google Sheet 環境變數與 Service Account 權限。'
        ].join('\\n'));
      }
    }
  }catch(err){
    console.error('[LINE_WEBHOOK_ERROR]', err);
  }
});

const server = app.listen(PORT, '0.0.0.0', () => console.log(`Noise dashboard server listening on 0.0.0.0:${PORT}`));
server.on('error', (err) => {
  console.error('[SERVER_LISTEN_ERROR]', err);
  process.exit(1);
});
