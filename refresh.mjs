// Cloud refresh: logs into Microsoft Graph, reads the 4 sales files, writes data.json (month-to-date)
import { writeFileSync, readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';

const { MS_TENANT_ID:TENANT, MS_CLIENT_ID:CLIENT, MS_CLIENT_SECRET:SECRET } = process.env;
const GRAPH = 'https://graph.microsoft.com/v1.0';

// If a rep renames a file / tab, update it here.
// basis: 'paid' = count only paid rows, bucket by the paid-date column (col with header "paid"), amount from "total paid" (fallback TOTAL).
// basis: 'paidflag' = count only rows whose "paid" column says PAID, bucket by order date, amount from TOTAL.
// no basis = all rows, bucket by order date, amount from TOTAL.
const REPS = [
  { rep:'Tehila',   user:'tcohen@surgamed.com', path:'Tehila/Tehila - Ventas.xlsx',          sheet:'Sheet1', basis:'paid' },
  { rep:'Deysi',    user:'dcalvo@surgamed.com', link:'https://netorgft3242511-my.sharepoint.com/:x:/g/personal/dcalvo_surgamed_com/IQB-BC_SvqB2RJhtHUCRkYp_AXgvOY4TjiU9bPaCwHmKkc4?e=se1JTU', path:'Deysi Sales list NEW COMPUTER.xlsx', sheet:'Deysi Sales', basis:'paid' },
  { rep:'Mirian',   user:'malejo@surgamed.com', path:'MIRI - PAYPAL SENT ORDERS.xlsx',        sheet:'Sheet1', basis:'paidflag' },
  { rep:'Jennifer', user:'jlugo@surgamed.com',  path:'Desktop/Jennifer Lugo Saless.xlsx',     sheet:'Invoiced-Sales', basis:'paid' },
];
const GOAL_MONTH = 200000; // adjust your team's MONTHLY revenue goal

async function getToken(){
  const body = new URLSearchParams({ client_id:CLIENT, client_secret:SECRET, grant_type:'client_credentials', scope:'https://graph.microsoft.com/.default' });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, { method:'POST', body });
  if(!r.ok) throw new Error('token '+r.status+' '+await r.text());
  return (await r.json()).access_token;
}
function shareId(url){ const b=Buffer.from(url,'utf8').toString('base64'); return 'u!'+b.replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-'); }
async function download(tok, R){
  let url;
  if(R.link){
    const sr = await fetch(`${GRAPH}/shares/${shareId(R.link)}/driveItem?$select=id,parentReference`, { headers:{ Authorization:`Bearer ${tok}` } });
    if(!sr.ok) throw new Error('share '+R.rep+' '+sr.status);
    const si = await sr.json();
    url = `${GRAPH}/drives/${si.parentReference.driveId}/items/${si.id}/content`;
  } else {
    const enc = R.path.split('/').map(encodeURIComponent).join('/');
    url = `${GRAPH}/users/${encodeURIComponent(R.user)}/drive/root:/${enc}:/content`;
  }
  const r = await fetch(url, { headers:{ Authorization:`Bearer ${tok}` } });
  if(!r.ok) throw new Error('download '+R.rep+' '+r.status);
  return Buffer.from(await r.arrayBuffer());
}
const EXCEL_EPOCH = Date.UTC(1899,11,30);
const toDate = v => { if(v==null||v==='') return null; if(v instanceof Date) return v;
  const n=Number(v); if(isFinite(n)&&n>20000&&n<60000) return new Date(EXCEL_EPOCH+Math.round(n)*86400000);
  const t=Date.parse(String(v)); return isNaN(t)?null:new Date(t); };
const num = v => v==null?NaN:parseFloat(String(v).replace(/[$,\s]/g,''));
const ymd = d => d.toISOString().slice(0,10);

function parseRep(buf, R){
  const wb = XLSX.read(buf, { cellDates:true });
  const ws = wb.Sheets[R.sheet] || wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:null });
  let hr=-1;
  for(let i=0;i<Math.min(15,grid.length);i++){
    const row=(grid[i]||[]).map(c=>String(c==null?'':c).trim().toLowerCase());
    if(row.includes('order date') && row.includes('total')){ hr=i; break; }
  }
  if(hr<0) return null;
  const H=(grid[hr]||[]).map(c=>String(c==null?'':c).trim().toLowerCase());
  const dc=H.indexOf('order date'), qc=H.findIndex(c=>c==='qty'||c==='quantity'), tc=H.indexOf('total'), nc=H.indexOf('name');
  const pc=H.indexOf('paid'), tpc=H.indexOf('total paid');
  const out=[];
  for(let i=hr+1;i<grid.length;i++){
    const row=grid[i]||[]; let d, amt;
    if(R.basis==='paid'){ d=toDate(row[pc]); if(!d) continue; amt=num(row[tpc]); if(!isFinite(amt)||amt<=0) amt=num(row[tc]); }
    else if(R.basis==='paidflag'){ const pf=String(row[pc]==null?'':row[pc]).toLowerCase(); if(!/paid|piad/.test(pf)) continue; d=toDate(row[dc]); amt=num(row[tc]); }
    else { d=toDate(row[dc]); amt=num(row[tc]); }
    if(!d || !isFinite(amt) || amt<=0) continue;
    const q=num(row[qc]);
    let nm=String(row[nc]==null?'':row[nc]).trim(); if(/[@]|http/i.test(nm)) nm=nm.split(/[@\s]/)[0]; nm=nm.slice(0,16);
    out.push({ rep:R.rep, name:nm, date:d, total:Math.round(amt*100)/100, qty:isFinite(q)?q:0 });
  }
  return out;
}

const tok = await getToken();
let all=[]; const files=[];
for(const R of REPS){
  try{ const buf=await download(tok,R); const o=parseRep(buf,R);
       if(o){ all=all.concat(o); files.push([R.rep,true]); } else files.push([R.rep,false]); }
  catch(e){ console.error(R.rep, e.message); files.push([R.rep,false]); }
}
const now=new Date();
const todayUTC=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()));
const monthStart=new Date(Date.UTC(todayUTC.getUTCFullYear(), todayUTC.getUTCMonth(), 1));
const monthRows=all.filter(o=>o.date>=monthStart && o.date<=todayUTC);
const byRep={};
for(const o of monthRows){ (byRep[o.rep]=byRep[o.rep]||{rev:0,ord:0,units:0}); byRep[o.rep].rev+=o.total; byRep[o.rep].ord++; byRep[o.rep].units+=o.qty; }
// Always list every rep whose file loaded (0 if no sales this month) so nobody drops off the board.
const okFiles = new Set(files.filter(f=>f[1]).map(f=>f[0]));
const reps=REPS.filter(R=>okFiles.has(R.rep)).map(R=>{ const v=byRep[R.rep]||{rev:0,ord:0,units:0}; return [R.rep, Math.round(v.rev*100)/100, v.ord, Math.round(v.units)]; }).sort((a,b)=>b[1]-a[1]);
const totals={revenue:0,orders:0,units:0};
reps.forEach(r=>{totals.revenue+=r[1];totals.orders+=r[2];totals.units+=r[3];});
totals.revenue=Math.round(totals.revenue*100)/100; totals.aov=totals.orders?Math.round(totals.revenue/totals.orders*100)/100:0;
const daysElapsed=todayUTC.getUTCDate();
const dailyRev=new Array(daysElapsed).fill(0), dailyOrd=new Array(daysElapsed).fill(0);
monthRows.forEach(o=>{ const d=o.date.getUTCDate()-1; if(d>=0&&d<daysElapsed){ dailyRev[d]+=o.total; dailyOrd[d]++; } });
const weekCut=new Date(todayUTC.getTime()-6*86400000);
const wk={};
all.filter(o=>o.date>=weekCut && o.date<=todayUTC).forEach(o=>{ wk[o.rep]=(wk[o.rep]||0)+o.total; });
const wkSorted=Object.entries(wk).sort((a,b)=>b[1]-a[1]);
const weekTop = wkSorted.length ? {rep:wkSorted[0][0], revenue:Math.round(wkSorted[0][1])} : {rep:'', revenue:0};
const biggest=monthRows.reduce((m,o)=>o.total>m.total?o:m,{total:0,rep:'',date:todayUTC});
const recent=[...monthRows].sort((a,b)=>b.date-a.date).slice(0,10).map(o=>[o.rep,o.name,ymd(o.date),o.total]);
const MON=['January','February','March','April','May','June','July','August','September','October','November','December'];
// Last month: per-rep totals, combined total, and same-days pace figure for a fair MTD comparison.
const lmStart=new Date(Date.UTC(monthStart.getUTCFullYear(), monthStart.getUTCMonth()-1, 1));
const lmRows=all.filter(o=>o.date>=lmStart && o.date<monthStart);
const lmBy={}; lmRows.forEach(o=>{ lmBy[o.rep]=(lmBy[o.rep]||0)+o.total; });
const lmSameBy={}; lmRows.filter(o=>o.date.getUTCDate()<=daysElapsed).forEach(o=>{ lmSameBy[o.rep]=(lmSameBy[o.rep]||0)+o.total; });
const lmReps=REPS.map(R=>[R.rep, Math.round((lmBy[R.rep]||0)*100)/100, Math.round((lmSameBy[R.rep]||0)*100)/100]).sort((a,b)=>b[1]-a[1]);
const lmTotal=Math.round(lmReps.reduce((s,r)=>s+r[1],0)*100)/100;
const lmSameDays=Math.round(lmReps.reduce((s,r)=>s+r[2],0)*100)/100;
const lastMonth={ name:MON[lmStart.getUTCMonth()], total:lmTotal, sameDays:lmSameDays, reps:lmReps };
// Trailing 3 full months per rep (oldest first) for the leaderboard mini-charts.
const hist3={ names:[], byRep:{} }; REPS.forEach(R=>hist3.byRep[R.rep]=[]);
for(let k=3;k>=1;k--){
  const s=new Date(Date.UTC(monthStart.getUTCFullYear(), monthStart.getUTCMonth()-k, 1));
  const e=new Date(Date.UTC(monthStart.getUTCFullYear(), monthStart.getUTCMonth()-k+1, 1));
  hist3.names.push(MON[s.getUTCMonth()].slice(0,3));
  const by={}; all.filter(o=>o.date>=s&&o.date<e).forEach(o=>{ by[o.rep]=(by[o.rep]||0)+o.total; });
  REPS.forEach(R=>hist3.byRep[R.rep].push(Math.round(by[R.rep]||0)));
}
const snap={ asOf:now.toISOString(), today:ymd(todayUTC), monthStart:ymd(monthStart), monthName:MON[todayUTC.getUTCMonth()], daysElapsed, goal:GOAL_MONTH,
  biggest:{rep:biggest.rep, amount:Math.round(biggest.total), date:ymd(biggest.date)},
  totals, reps, weekTop, dailyRev:dailyRev.map(x=>Math.round(x*100)/100), dailyOrd, recent, files, lastMonth, hist3 };
writeFileSync('data.json', JSON.stringify(snap));
let pageHtml = readFileSync('index.html','utf8');
pageHtml = pageHtml.replace(/(<script id="snapshot-data"[^>]*>)[\s\S]*?(<\/script>)/, (m,a,b)=> a + '\n' + JSON.stringify(snap) + '\n' + b);
// SurgaMed cross-app nav — survives every refresh
if (!pageHtml.includes("smNav")) {
  const __smNav = "<script>(function(){function add(){if(document.getElementById(\"smNav\")||!document.body)return;var n=document.createElement(\"div\");n.id=\"smNav\";n.style.cssText=\"position:fixed;top:8px;right:10px;z-index:9999;display:flex;gap:6px;font:12px -apple-system,Segoe UI,Roboto,sans-serif;opacity:.85\";n.innerHTML='<a href=\"leads.html\" style=\"background:#185fa5;color:#fff;padding:4px 10px;border-radius:999px;text-decoration:none;font-weight:600\">\\uD83D\\uDCCB Leads</a><a href=\"whatsapp-crm.html\" style=\"background:#185fa5;color:#fff;padding:4px 10px;border-radius:999px;text-decoration:none;font-weight:600\">\\uD83D\\uDCAC CRM</a>';document.body.appendChild(n);}add();document.addEventListener(\"DOMContentLoaded\",add);setTimeout(add,1500);})();</script>";
  pageHtml = pageHtml.includes("</body>") ? pageHtml.replace("</body>", __smNav + "</body>") : pageHtml + __smNav;
}
writeFileSync('index.html', pageHtml);
console.log('OK', JSON.stringify({reps, totals, files}));
