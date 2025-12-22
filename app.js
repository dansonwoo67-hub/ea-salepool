
const $ = (id) => document.getElementById(id);

let wb1=null, wb2=null;
let t1=null, t2=null; // {rows, dateMax}
let mapping=null;

const POOL_EXCLUDE = new Set(["m1"]);
const CANON = {
  learnerId:"learnerId", eaName:"eaName", teamName:"teamName",
  lastConn:"lastConn", lastFollow:"lastFollow", remaining:"remaining", lastMonthConsumed:"lastMonthConsumed"
};

const ALIASES = {
  learnerId: ["learner id","learner_id","student id","student_id","id"],
  eaName: ["latest (current) assigned ss staff name","assigned ss staff name","ss staff name","staff name","ea name","ss name","name"],
  teamName: ["latest (current) assigned ss group name","assigned ss group name","ss group name","group name","team","team name"],
  lastConn: ["last ss connection date","last connection date","last connect date","last connected date","connection date"],
  lastFollow: ["last follow up date","last follow-up date","last ss follow up date","last ss follow-up date","follow up date","follow-up date","last touch date","last contacted date"],
  remaining: ["remaining lessons","remaining class","remaining classes","remaining","left lessons","left class"],
  lastMonthConsumed: ["last month consumed","last month consumption","last month used","last month usage","lm consumed","consumed last month","lastmonth consumed","last month"]
};

function norm(s){ return (s??"").toString().trim().toLowerCase(); }

function poolNameFromSheet(sheetName){
  const n = norm(sheetName);
  if (n.includes("m2")) return "m2";
  if (n.includes("expiring")) return "expiring";
  if (n.includes("expired")) return "expired";
  if (n.includes("duration")) return "duration";
  if (n.includes("period")) return "period";
  if (n === "exp" || n.includes(" exp")) return "exp";
  if (n.includes("m1")) return "m1";
  return n;
}

function excelDateToISO(v){
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  if (typeof v === "number" && isFinite(v)){
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    const dt = new Date(Date.UTC(d.y, d.m-1, d.d));
    return dt.toISOString().slice(0,10);
  }
  const s = v.toString().trim();
  if (s.startsWith("1970") || s.startsWith("1900")) return null;
  const ms = Date.parse(s);
  if (!isNaN(ms)) return new Date(ms).toISOString().slice(0,10);
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m1){
    let dd=+m1[1], mm=+m1[2], yy=+m1[3]; if (yy<100) yy+=2000;
    const dt = new Date(Date.UTC(yy, mm-1, dd));
    if (!isNaN(dt)) return dt.toISOString().slice(0,10);
  }
  return null;
}
function isValidISODate(iso){
  if (!iso) return false;
  if (iso.startsWith("1970") || iso.startsWith("1900")) return false;
  return /^\d{4}-\d{2}-\d{2}$/.test(iso);
}
function safeNum(v){ const n=Number(v); return isNaN(n)?null:n; }

function uniqueHeadersFromWorkbook(wb){
  const set=new Set();
  wb.SheetNames.forEach(sn=>{
    const ws=wb.Sheets[sn];
    const arr = XLSX.utils.sheet_to_json(ws,{header:1, defval:"", blankrows:false});
    if (arr.length>0){ (arr[0]||[]).forEach(h=>set.add(h)); }
  });
  return Array.from(set).filter(x=>x!=="" && x!=null);
}
function detectColumns(headers){
  const map={};
  const hn=headers.map(h=>({raw:h, n:norm(h)}));
  for (const key of Object.keys(ALIASES)){
    let found=null;
    for (const a of ALIASES[key]){
      const an=norm(a);
      const hit=hn.find(x=> x.n===an || x.n.includes(an));
      if (hit){ found=hit.raw; break; }
    }
    map[key]=found||null;
  }
  return map;
}
function buildMappingUI(allHeaders, autoMap){
  const area=$("mappingArea"); area.innerHTML="";
  const fields=[
    [CANON.learnerId,"Learner ID *"],
    [CANON.eaName,"EA Name *"],
    [CANON.teamName,"Team Name"],
    [CANON.lastConn,"Last Connection Date *"],
    [CANON.lastFollow,"Last Follow-up Date (optional)"],
    [CANON.remaining,"Remaining Lessons (optional)"],
    [CANON.lastMonthConsumed,"Last Month Consumed (optional)"],
  ];
  for (const [key,label] of fields){
    const div=document.createElement("div"); div.className="mapItem";
    const lab=document.createElement("label"); lab.textContent=label;
    const sel=document.createElement("select"); sel.id=`map_${key}`;
    const opt0=document.createElement("option"); opt0.value=""; opt0.textContent="(not used)";
    sel.appendChild(opt0);
    for (const h of allHeaders){
      const opt=document.createElement("option"); opt.value=h; opt.textContent=h; sel.appendChild(opt);
    }
    if (autoMap && autoMap[key]) sel.value=autoMap[key];
    div.appendChild(lab); div.appendChild(sel);
    area.appendChild(div);
  }
}

function parseWorkbookFast(wb, chosenMap, teamFilter){
  const rows=[];
  let dateMax=null;

  wb.SheetNames.forEach(sn=>{
    const pool=poolNameFromSheet(sn);
    const ws=wb.Sheets[sn];
    const aoa = XLSX.utils.sheet_to_json(ws,{header:1, defval:"", blankrows:false, raw:true});
    if (!aoa || aoa.length<2) return;
    const header = aoa[0].map(x=>x==null?"":x.toString());
    const idx = {};
    for (const key of Object.values(CANON)){
      const colName = chosenMap[key];
      if (!colName) { idx[key] = -1; continue; }
      idx[key] = header.findIndex(h => h === colName);
    }

    for (let i=1;i<aoa.length;i++){
      const r=aoa[i];
      const learnerId = idx.learnerId>=0 ? r[idx.learnerId] : null;
      const eaName    = idx.eaName>=0 ? r[idx.eaName] : null;
      if (!learnerId || !eaName) continue;

      const teamName  = idx.teamName>=0 ? r[idx.teamName] : "";
      if (teamFilter && teamName && teamName.toString().trim() !== teamFilter.trim()) continue;

      const lastConn  = idx.lastConn>=0 ? excelDateToISO(r[idx.lastConn]) : null;
      const lastFollow= idx.lastFollow>=0 ? excelDateToISO(r[idx.lastFollow]) : null;
      const remaining = idx.remaining>=0 ? safeNum(r[idx.remaining]) : null;
      const lmCons    = idx.lastMonthConsumed>=0 ? safeNum(r[idx.lastMonthConsumed]) : null;

      if (isValidISODate(lastConn)){
        if (!dateMax || lastConn > dateMax) dateMax = lastConn;
      }

      rows.push({
        pool,
        learnerId: learnerId.toString().trim(),
        eaName: eaName.toString().trim(),
        teamName: (teamName??"").toString().trim(),
        lastConn,
        lastFollow,
        remaining,
        lastMonthConsumed: lmCons
      });
    }
  });

  return {rows, dateMax};
}

function isCovered(r){ return isValidISODate(r.lastConn); }

function monthStartISO(iso){ const [y,m]=iso.split("-").map(Number); return new Date(Date.UTC(y,m-1,1)).toISOString().slice(0,10); }

function isTouchedThisMonth(r, reportDate){
  const ms = monthStartISO(reportDate);
  const touch = r.lastFollow || r.lastConn;
  if (!isValidISODate(touch)) return false;
  return touch >= ms && touch <= reportDate;
}
function k(parts){ return parts.map(x=>(x??"").toString()).join("||"); }

function aggCoverage(rows){
  const members=new Map(), covered=new Map();
  for (const r of rows){
    const key=k([r.teamName,r.eaName,r.pool]);
    if (!members.has(key)) members.set(key,new Set());
    members.get(key).add(r.learnerId);
    if (isCovered(r)){
      if (!covered.has(key)) covered.set(key,new Set());
      covered.get(key).add(r.learnerId);
    }
  }
  const out=[];
  for (const [key,ms] of members){
    const [team,ea,pool]=key.split("||");
    const m=ms.size, c=(covered.get(key)?.size)||0;
    out.push({team,ea,pool,members:m,covered:c,rate:m?c/m:0});
  }
  out.sort((a,b)=>(a.team+a.ea+a.pool).localeCompare(b.team+b.ea+b.pool));
  return out;
}

function deltaCoverage(a1,a2){
  const m1=new Map(a1.map(x=>[k([x.team,x.ea,x.pool]),x]));
  const m2=new Map(a2.map(x=>[k([x.team,x.ea,x.pool]),x]));
  const keys=new Set([...m1.keys(),...m2.keys()]);
  const out=[];
  keys.forEach(key=>{
    const [team,ea,pool]=key.split("||");
    const x1=m1.get(key)||{members:0,covered:0,rate:0};
    const x2=m2.get(key)||{members:0,covered:0,rate:0};
    out.push({
      team,ea,pool,
      members_t1:x1.members, covered_t1:x1.covered, rate_t1:(x1.rate*100).toFixed(1)+"%",
      members_t2:x2.members, covered_t2:x2.covered, rate_t2:(x2.rate*100).toFixed(1)+"%",
      covered_delta:(x2.covered-x1.covered),
      rate_delta:((x2.rate-x1.rate)*100).toFixed(1)+"%"
    });
  });
  out.sort((a,b)=>(a.team+a.ea+a.pool).localeCompare(b.team+b.ea+b.pool));
  return out;
}

function identifyFirstFollow(rows1, rows2, t2Date){
  const idx1=new Map();
  rows1.forEach(r=>idx1.set(k([r.teamName,r.eaName,r.pool,r.learnerId]),r));
  const idx2=new Map();
  rows2.forEach(r=>idx2.set(k([r.teamName,r.eaName,r.pool,r.learnerId]),r));
  const keys=new Set([...idx1.keys(),...idx2.keys()]);
  const agg=new Map();

  keys.forEach(key=>{
    const r2=idx2.get(key); if (!r2) return;
    const r1=idx1.get(key)||null;
    const c1=r1?isCovered(r1):false;
    const c2=isCovered(r2);
    if (!c2) return;

    let type=null;
    if (!c1 && c2) type="FIRST_ROUND";
    else if (c1 && c2){
      const lc1=r1.lastConn||"", lc2=r2.lastConn||"";
      if (lc2 && (!lc1 || lc2>lc1)){
        type = (t2Date && lc2===t2Date) ? "FOLLOW_UP_ON_LATEST" : "FOLLOW_UP_UPDATED";
      }
    }
    if (!type) return;
    const kk=k([r2.teamName,r2.eaName,r2.pool,type]);
    agg.set(kk,(agg.get(kk)||0)+1);
  });

  const out=[];
  agg.forEach((cnt,kk)=>{
    const [team,ea,pool,type]=kk.split("||");
    out.push({team,ea,pool,type,count:cnt});
  });
  out.sort((a,b)=>(a.team+a.ea+a.pool+a.type).localeCompare(b.team+b.ea+b.pool+b.type));
  return out;
}

function buildRecommendations(rows2, reportDate){
  const ms = monthStartISO(reportDate);

  const notCoveredThisMonth = (r)=> !isCovered(r) || (r.lastConn < ms);
  const notTouchedThisMonth = (r)=> !isTouchedThisMonth(r, reportDate);

  const scoreExp = (r)=> (r.lastMonthConsumed ?? 0) + (notTouchedThisMonth(r)?50:0);
  const scorePD  = (r)=>{
    const lm=r.lastMonthConsumed??0;
    const rem=(r.remaining==null?999999:r.remaining);
    const remScore = rem<=0?100:(100/(rem+1));
    return remScore + 0.8*lm + (notTouchedThisMonth(r)?50:0);
  };

  const byEA=new Map();
  for (const r of rows2){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    if (!byEA.has(r.eaName)) byEA.set(r.eaName,[]);
    byEA.get(r.eaName).push(r);
  }

  const rec=[];
  byEA.forEach((arr)=>{
    const picked=new Set();
    const push=(r,reason,score)=>{
      const key=r.learnerId+"||"+r.pool;
      if (picked.has(key)) return false;
      picked.add(key);
      rec.push({
        team:r.teamName, ea:r.eaName, pool:r.pool, learnerId:r.learnerId,
        output:`${r.eaName} - ${r.learnerId} - ${r.pool}`,
        reason, score:Number(score.toFixed(3))
      });
      return true;
    };

    const m2=arr.filter(r=>r.pool==="m2" && notCoveredThisMonth(r))
      .sort((a,b)=>(a.lastConn||"").localeCompare(b.lastConn||""));
    const expiring=arr.filter(r=>r.pool==="expiring" && notTouchedThisMonth(r))
      .sort((a,b)=>scoreExp(b)-scoreExp(a));
    const expired=arr.filter(r=>r.pool==="expired" && notTouchedThisMonth(r))
      .sort((a,b)=>scoreExp(b)-scoreExp(a));
    const pd=arr.filter(r=>(r.pool==="period"||r.pool==="duration") && notTouchedThisMonth(r))
      .sort((a,b)=>scorePD(b)-scorePD(a));

    let n=0;
    for (const r of m2){ if (n>=20) break; if (push(r,"M2: not covered this month",1000)) n++; }
    for (const r of expiring){ if (n>=20) break; if (push(r,"Expiring: high last-month consumption & not touched this month",scoreExp(r))) n++; }
    for (const r of expired){ if (n>=20) break; if (push(r,"Expired: high last-month consumption & not touched this month",scoreExp(r))) n++; }
    for (const r of pd){ if (n>=20) break; if (push(r,"Period/Duration: low remaining + high consumption & not touched this month",scorePD(r))) n++; }
  });

  rec.sort((a,b)=>{
    if (a.team!==b.team) return a.team.localeCompare(b.team);
    if (a.ea!==b.ea) return a.ea.localeCompare(b.ea);
    return b.score-a.score;
  });
  return rec;
}

function renderTable(containerId, columns, rows){
  const wrap=$(containerId);
  if (!rows || rows.length===0){ wrap.innerHTML=""; return; }
  let html="<table><thead><tr>";
  for (const c of columns) html+=`<th>${c}</th>`;
  html+="</tr></thead><tbody>";
  for (const r of rows){
    html+="<tr>";
    for (const c of columns) html+=`<td>${r[c] ?? ""}</td>`;
    html+="</tr>";
  }
  html+="</tbody></table>";
  wrap.innerHTML=html;
}
function downloadCSV(filename, rows){
  if (!rows || rows.length===0) return;
  const cols=Object.keys(rows[0]);
  const esc=(s)=>`"${(s??"").toString().replaceAll('"','""')}"`;
  const lines=[cols.map(esc).join(",")];
  for (const r of rows) lines.push(cols.map(c=>esc(r[c])).join(","));
  const blob=new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
}

async function readWorkbook(file){
  const buf=await file.arrayBuffer();
  return XLSX.read(buf,{type:"array", cellDates:true, dense:true});
}

function setupTabs(){
  document.querySelectorAll(".tab").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
      document.querySelectorAll(".tabpane").forEach(x=>x.classList.remove("active"));
      btn.classList.add("active");
      $(btn.dataset.tab).classList.add("active");
    });
  });
}
setupTabs();

$("btnLoad").addEventListener("click", async ()=>{
  const f1=$("file1").files[0], f2=$("file2").files[0];
  const teamFilter=$("teamFilter").value || "";
  if (!f1 || !f2) return alert("Please select both T1 and T2 files.");

  $("t1meta").textContent="Loading...";
  $("t2meta").textContent="Loading...";

  wb1 = await readWorkbook(f1);
  wb2 = await readWorkbook(f2);

  const headers=[...new Set([...uniqueHeadersFromWorkbook(wb1), ...uniqueHeadersFromWorkbook(wb2)])];
  const autoMap=detectColumns(headers);
  buildMappingUI(headers, autoMap);
  $("mappingCard").style.display="block";

  $("t1meta").textContent=`Sheets: ${wb1.SheetNames.length}`;
  $("t2meta").textContent=`Sheets: ${wb2.SheetNames.length}`;

  if (autoMap.learnerId && autoMap.eaName && autoMap.lastConn){
    mapping=autoMap;
    t1=parseWorkbookFast(wb1, mapping, teamFilter);
    t2=parseWorkbookFast(wb2, mapping, teamFilter);
    $("t1meta").textContent+=` | Latest contact: ${t1.dateMax || "-"}`;
    $("t2meta").textContent+=` | Latest contact: ${t2.dateMax || "-"}`;
    $("btnAnalyze").disabled=false;
  } else {
    $("btnAnalyze").disabled=true;
    alert("Auto-detection failed. Please map Learner ID, EA Name, and Last Connection Date.");
  }
});

$("btnApplyMapping").addEventListener("click", ()=>{
  if (!wb1 || !wb2) return alert("Load files first.");
  const teamFilter=$("teamFilter").value || "";

  const chosen={};
  for (const key of Object.values(CANON)){
    const sel=document.getElementById(`map_${key}`);
    chosen[key]=sel?(sel.value||null):null;
  }
  if (!chosen.learnerId || !chosen.eaName || !chosen.lastConn){
    return alert("Learner ID, EA Name, and Last Connection Date are required.");
  }
  mapping=chosen;
  t1=parseWorkbookFast(wb1, mapping, teamFilter);
  t2=parseWorkbookFast(wb2, mapping, teamFilter);
  $("t1meta").textContent=`Sheets: ${wb1.SheetNames.length} | Latest contact: ${t1.dateMax || "-"}`;
  $("t2meta").textContent=`Sheets: ${wb2.SheetNames.length} | Latest contact: ${t2.dateMax || "-"}`;
  $("btnAnalyze").disabled=false;
});

$("btnAnalyze").addEventListener("click", ()=>{
  if (!t1 || !t2) return alert("Load files first.");
  const reportDate2 = t2.dateMax || t1.dateMax;
  if (!reportDate2) return alert("Cannot find any valid 'Last Connection Date' in the file(s).");

  $("resultsCard").style.display="block";

  const cov1=aggCoverage(t1.rows);
  const cov2=aggCoverage(t2.rows);
  const delta=deltaCoverage(cov1,cov2);
  renderTable("coverageTable",
    ["team","ea","pool","members_t1","covered_t1","rate_t1","members_t2","covered_t2","rate_t2","covered_delta","rate_delta"],
    delta
  );

  const ff=identifyFirstFollow(t1.rows,t2.rows, reportDate2);
  renderTable("firstFollowTable", ["team","ea","pool","type","count"], ff);

  const rec=buildRecommendations(t2.rows, reportDate2);
  renderTable("recommendTable", ["team","ea","pool","learnerId","output","reason","score"], rec);

  $("exportCoverage").onclick=()=>downloadCSV(`coverage_delta.csv`, delta);
  $("exportFirstFollow").onclick=()=>downloadCSV(`first_follow.csv`, ff);
  $("exportRecommend").onclick=()=>downloadCSV(`recommend.csv`, rec);
});
