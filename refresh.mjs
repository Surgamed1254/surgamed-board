// Cloud refresh: logs into Microsoft Graph, reads the 4 sales files, writes data.json
import { writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';

const { MS_TENANT_ID:TENANT, MS_CLIENT_ID:CLIENT, MS_CLIENT_SECRET:SECRET } = process.env;
const GRAPH = 'https://graph.microsoft.com/v1.0';

// If a rep renames a file / tab, update it here.
const REPS = [
  { rep:'Tehila',   user:'tcohen@surgamed.com', path:'Tehila/Tehila - Ventas.xlsx',          sheet:'Sheet1' },
  { rep:'Deysi',    user:'dcalvo@surgamed.com', path:'Deysi Sales list NEW COMPUTER.xlsx',    sheet:'Deysi Sales' },
  { rep:'Miri',     user:'malejo@surgamed.com', path:'MIRI - PAYPAL SENT ORDERS.xlsx',        sheet:'Sheet1' },
  { rep:'Jennifer', user:'jlugo@surgamed.com',  path:'Desktop/Jennifer Lugo Saless.xlsx',     sheet:'Invoiced-Sales' },
];
const GOAL_30D = 150000; // adjust your team's 30-day revenue goal

async function getToken(){
  const body = new URLSearchParams({ client_id:CLIENT, client_secret:SECRET, grant_type:'client_credentials', scope:'https://graph.microsoft.com/.default' });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, { method:'POST', body });
  if(!r.ok) throw new Error('token '+r.status+' '+await r.text());
  return (await r.json()).access_token;
}
async function download(tok, user, path){
  const enc = path.split('/').map(encodeURIComponent).join('/');
  const r = await fetch(`${GRAPH}/users/${encodeURIComponent(user)}/drive/root:/${enc}:/content`, { headers:{ Authorization:`Bearer ${tok}` } });
  if(!r.ok) throw new Error('download '+user+' '+r.status);
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
  const out=[];
  for(let i=hr+1;i<grid.length;i++){
    const row=grid[i]||[]; const d=toDate(row[dc]); if(!d) continue;
    const tot=num(row[tc]); if(!isFinite(tot)||tot<=0) continue;
    const q=num(row[qc]);
    let nm=String(row[nc]==null?'':row[nc]).trim(); if(/[@]|http/i.test(nm)) nm=nm.split(/[@\s]/)[0]; nm=nm.slice(0,16);
    out.push({ rep:R.rep, name:nm, date:d, total:Math.round(tot*100)/100, qty:isFinite(q)?q:0 });
  }
  return out;
}

const tok = await getToken();
let all=[]; const files=[];
for(const R of REPS){
  try{ const buf=await download(tok,R.user,R.path); const o=parseRep(buf,R);
       if(o){ all=all.concat(o); files.push([R.rep,true]); } else files.push([R.rep,false]); }
  catch(e){ console.error(R.rep, e.message); files.push([R.rep,false]); }
}
const now=new Date();
const todayUTC=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()));
const cutoff=new Date(todayUTC.getTime()-29*86400000);
const win=all.filter(o=>o.date>=cutoff && o.date<=todayUTC);
const monthStart=new Date(Date.UTC(todayUTC.getUTCFullYear(), todayUTC.getUTCMonth(), 1));
const mtdRows=all.filter(o=>o.date>=monthStart && o.date<=todayUTC);
const mtd={ revenue:Math.round(mtdRows.reduce((a,o)=>a+o.total,0)), orders:mtdRows.length, units:Math.round(mtdRows.reduce((a,o)=>a+o.qty,0)) };
const byRep={};
for(const o of win){ (byRep[o.rep]=byRep[o.rep]||{rev:0,ord:0,units:0}); byRep[o.rep].rev+=o.total; byRep[o.rep].ord++; byRep[o.rep].units+=o.qty; }
const reps=Object.entries(byRep).map(([r,v])=>[r,Math.round(v.rev*100)/100,v.ord,Math.round(v.units)]).sort((a,b)=>b[1]-a[1]);
const totals={revenue:0,orders:0,units:0};
reps.forEach(r=>{totals.revenue+=r[1];totals.orders+=r[2];totals.units+=r[3];});
totals.revenue=Math.round(totals.revenue*100)/100; totals.aov=totals.orders?Math.round(totals.revenue/totals.orders*100)/100:0;
const dailyRev=[],dailyOrd=[],idx={};
for(let i=0;i<30;i++){ const k=ymd(new Date(cutoff.getTime()+i*86400000)); idx[k]=i; dailyRev[i]=0; dailyOrd[i]=0; }
win.forEach(o=>{ const k=ymd(o.date); if(idx[k]!=null){ dailyRev[idx[k]]+=o.total; dailyOrd[idx[k]]++; } });
const weekCut=new Date(todayUTC.getTime()-6*86400000);
const wk={};
all.filter(o=>o.date>=weekCut && o.date<=todayUTC).forEach(o=>{ wk[o.rep]=(wk[o.rep]||0)+o.total; });
const wkSorted=Object.entries(wk).sort((a,b)=>b[1]-a[1]);
const weekTop = wkSorted.length ? {rep:wkSorted[0][0], revenue:Math.round(wkSorted[0][1])} : {rep:'', revenue:0};
const biggest=win.reduce((m,o)=>o.total>m.total?o:m,{total:0,rep:'',date:todayUTC});
const recent=[...win].sort((a,b)=>b.date-a.date).slice(0,10).map(o=>[o.rep,o.name,ymd(o.date),o.total]);
const snap={ asOf:now.toISOString(), cutoff:ymd(cutoff), today:ymd(todayUTC), goal30d:GOAL_30D,
  biggest:{rep:biggest.rep, amount:Math.round(biggest.total), date:ymd(biggest.date)},
  totals, mtd, reps, weekTop, dailyRev:dailyRev.map(x=>Math.round(x*100)/100), dailyOrd, recent, files };
writeFileSync('data.json', JSON.stringify(snap));
console.log('OK', JSON.stringify({reps, totals, files}));
