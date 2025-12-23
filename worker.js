
importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

const POOL_EXCLUDE = new Set(["m1"]);
const FIXED = {
  learnerId: "Learner ID",
  familyId: "Family ID",
  eaName: "Latest (Current) Assigned SS Staff Name",
  teamName: "Latest (Current) Assigned SS Group Name",
  lastConn: "Last SS Connection Date",
  lastMonthConsumed: "Class Consumption Last Month",
  thisMonthConsumed: "Class Consumption This Month",
  remaining: "Total Session Card Count"
};

function norm(s){ return (s??"").toString().trim().toLowerCase(); }
function poolNameFromSheet(sheetName){
  const n = norm(sheetName);
  if (n.includes("m2")) return "m2";
  if (n.includes("expiring")) return "expiring";
  if (n.includes("duration")) return "duration";
  if (n.includes("expired")) return "expired";
  if (n === "exp" || n.startsWith("exp ")) return "expired";
  if (n.includes("period")) return "period";
  if (n.includes("m1")) return "m1";
  return n;
}

function isoFromYMD(y,m,d){
  const yy = String(y).padStart(4,"0");
  const mm = String(m).padStart(2,"0");
  const dd = String(d).padStart(2,"0");
  return `${yy}-${mm}-${dd}`;
}
function addDays(iso, days){
  const y=Number(iso.slice(0,4)), m=Number(iso.slice(5,7)), d=Number(iso.slice(8,10));
  const dt = new Date(Date.UTC(y,m-1,d));
  dt.setUTCDate(dt.getUTCDate()+days);
  return isoFromYMD(dt.getUTCFullYear(), dt.getUTCMonth()+1, dt.getUTCDate());
}
function lastDayOfMonth(y,m){
  const dt = new Date(Date.UTC(y, m, 0));
  return isoFromYMD(dt.getUTCFullYear(), dt.getUTCMonth()+1, dt.getUTCDate());
}
function diffDays(aIso, bIso){
  const ay=Number(aIso.slice(0,4)), am=Number(aIso.slice(5,7)), ad=Number(aIso.slice(8,10));
  const by=Number(bIso.slice(0,4)), bm=Number(bIso.slice(5,7)), bd=Number(bIso.slice(8,10));
  const a=new Date(Date.UTC(ay,am-1,ad));
  const b=new Date(Date.UTC(by,bm-1,bd));
  return Math.floor((a-b)/86400000);
}

function excelDateToISO(v){
  if (v == null || v === "") return null;

  if (typeof v === "number" && isFinite(v)){
    const d = XLSX.SSF.parse_date_code(v);
    if (!d || !d.y || !d.m || !d.d) return null;
    return isoFromYMD(d.y, d.m, d.d);
  }
  if (v instanceof Date && !isNaN(v)){
    return isoFromYMD(v.getUTCFullYear(), v.getUTCMonth()+1, v.getUTCDate());
  }

  const s = v.toString().trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:\s|T|$)/);
  if (m){
    const y=Number(m[1]), mo=Number(m[2]), d=Number(m[3]);
    if (y>=2000 && mo>=1 && mo<=12 && d>=1 && d<=31) return isoFromYMD(y,mo,d);
  }

  m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:\s|T|$)/);
  if (m){
    const d=Number(m[1]), mo=Number(m[2]), y=Number(m[3]);
    if (y>=2000 && mo>=1 && mo<=12 && d>=1 && d<=31) return isoFromYMD(y,mo,d);
  }

  const ms = Date.parse(s);
  if (!isNaN(ms)){
    const dt = new Date(ms);
    return isoFromYMD(dt.getUTCFullYear(), dt.getUTCMonth()+1, dt.getUTCDate());
  }
  return null;
}

function isValidISODate(iso){
  if (!iso) return false;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
  const y = Number(iso.slice(0,4));
  if (y < 2000) return false;
  return true;
}

function safeNum(v){ const n=Number(v); return isNaN(n)?null:n; }

function readWorkbookFast(buf){
  return XLSX.read(buf, {
    type:"array",
    dense:true,
    cellDates:false,
    cellNF:false,
    cellText:false,
    cellStyles:false,
    sheetStubs:false
  });
}

function headersContain(wb, col){
  for (const sn of wb.SheetNames){
    const ws = wb.Sheets[sn];
    const aoa = XLSX.utils.sheet_to_json(ws,{header:1, defval:"", blankrows:false, raw:true});
    if (!aoa || aoa.length<1) continue;
    const header = (aoa[0]||[]).map(x=>x==null?"":x.toString());
    if (header.includes(col)) return true;
  }
  return false;
}
function fixedTemplateWorks(wb){
  return headersContain(wb, FIXED.learnerId) && headersContain(wb, FIXED.eaName) && headersContain(wb, FIXED.lastConn);
}

function parseWorkbook(wb, teamFilter){
  const rows=[];
  for (const sn of wb.SheetNames){
    const pool=poolNameFromSheet(sn);
    const ws=wb.Sheets[sn];
    const aoa = XLSX.utils.sheet_to_json(ws,{header:1, defval:"", blankrows:false, raw:true});
    if (!aoa || aoa.length<2) continue;

    const header = aoa[0].map(x=>x==null?"":x.toString());
    const idx = {
      learnerId: header.indexOf(FIXED.learnerId),
      familyId: header.indexOf(FIXED.familyId),
      eaName: header.indexOf(FIXED.eaName),
      teamName: header.indexOf(FIXED.teamName),
      lastConn: header.indexOf(FIXED.lastConn),
      lastMonthConsumed: header.indexOf(FIXED.lastMonthConsumed),
      thisMonthConsumed: header.indexOf(FIXED.thisMonthConsumed),
      remaining: header.indexOf(FIXED.remaining)
    };
    if (idx.learnerId<0 || idx.eaName<0 || idx.lastConn<0) continue;

    for (let i=1;i<aoa.length;i++){
      const r=aoa[i];
      const learnerId = r[idx.learnerId];
      const eaName = r[idx.eaName];
      if (!learnerId || !eaName) continue;

      const teamName = idx.teamName>=0 ? r[idx.teamName] : "";
      if (teamFilter && teamName && teamName.toString().trim() !== teamFilter.trim()) continue;

      const lastConn = excelDateToISO(idx.lastConn>=0 ? r[idx.lastConn] : null);
      const familyId = idx.familyId>=0 ? (r[idx.familyId] ?? "") : "";
      const lm = idx.lastMonthConsumed>=0 ? safeNum(r[idx.lastMonthConsumed]) : null;
      const tm = idx.thisMonthConsumed>=0 ? safeNum(r[idx.thisMonthConsumed]) : null;
      const rem = idx.remaining>=0 ? safeNum(r[idx.remaining]) : null;

      rows.push({
        pool,
        learnerId: learnerId.toString().trim(),
        familyId: (familyId??"").toString().trim(),
        eaName: eaName.toString().trim(),
        teamName: (teamName??"").toString().trim(),
        lastConn,
        lastMonthConsumed: lm,
        thisMonthConsumed: tm,
        remaining: rem
      });
    }
  }
  return rows;
}

function monthBoundsFromPicker(monthPick, fallbackRows){
  let y, m;
  if (monthPick && /^\d{4}-\d{2}$/.test(monthPick)){
    y = Number(monthPick.slice(0,4));
    m = Number(monthPick.slice(5,7));
  } else {
    let max=null;
    for (const r of fallbackRows){
      if (isValidISODate(r.lastConn) && (!max || r.lastConn>max)) max=r.lastConn;
    }
    if (!max){
      const dt=new Date();
      y=dt.getUTCFullYear(); m=dt.getUTCMonth()+1;
    } else {
      y = Number(max.slice(0,4));
      m = Number(max.slice(5,7));
    }
  }
  const label = `${String(y).padStart(4,"0")}-${String(m).padStart(2,"0")}`;
  return {start: isoFromYMD(y,m,1), endFull: lastDayOfMonth(y,m), label};
}

function maxDateInRange(rows, start, end){
  let max=null;
  for (const r of rows){
    const d=r.lastConn;
    if (isValidISODate(d) && d>=start && d<=end){
      if (!max || d>max) max=d;
    }
  }
  return max;
}

function inRange(d, start, end){
  return isValidISODate(d) && d>=start && d<=end;
}

function k(parts){ return parts.map(x=>(x??"").toString()).join("||"); }

function poolPriority(pool){
  if (pool==="m2") return 4;
  if (pool==="expiring") return 3;
  if (pool==="duration" || pool==="period") return 2;
  if (pool==="expired") return 1;
  return 0;
}

function buildMetricsAndRemark(t1Rows, t2Rows, monthStart, monthEndFull, t1End, t2End){
  const t1Map=new Map();
  for (const r of t1Rows){
    const key=k([r.teamName,r.eaName,r.pool,r.learnerId]);
    const cur=t1Map.get(key);
    if (!cur || (r.lastConn||"")>(cur.lastConn||"")) t1Map.set(key, r);
  }
  const t2Map=new Map();
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    const key=k([r.teamName,r.eaName,r.pool,r.learnerId]);
    const cur=t2Map.get(key);
    if (!cur || (r.lastConn||"")>(cur.lastConn||"")) t2Map.set(key, r);
  }

  const poolAgg=new Map();
  const eaAgg=new Map();
  const poolDeltaByEA=new Map(); // eaKey -> pool -> delta

  const periodStart = t1End ? t1End : addDays(monthStart, -1);
  const periodStartNext = addDays(periodStart, 1);

  for (const [key, r2] of t2Map.entries()){
    const [team,ea,pool] = key.split("||").slice(0,3);
    const poolKey=k([team,pool,ea]);
    const eaKey=k([team,ea]);

    if (!poolAgg.has(poolKey)){
      poolAgg.set(poolKey,{team,pool,ea,totalSet:new Set(), t1MonthSet:new Set(), t2MonthSet:new Set(), pFirstSet:new Set(), pFollowSet:new Set()});
    }
    if (!eaAgg.has(eaKey)){
      eaAgg.set(eaKey,{team,ea,total:0,t1Month:0,t2Month:0,pFirst:0,pFollow:0});
    }
    if (!poolDeltaByEA.has(eaKey)) poolDeltaByEA.set(eaKey, new Map());

    const pa=poolAgg.get(poolKey);
    pa.totalSet.add(r2.learnerId);

    const r1=t1Map.get(key);
    const d1=r1 ? r1.lastConn : null;
    const d2=r2.lastConn;

    const t1Covered = t1End ? inRange(d1, monthStart, t1End) : false;
    const t2Covered = t2End ? inRange(d2, monthStart, t2End) : false;

    if (t1Covered) pa.t1MonthSet.add(r2.learnerId);
    if (t2Covered) pa.t2MonthSet.add(r2.learnerId);

    if (t2End && inRange(d2, periodStartNext, t2End) && t2Covered){
      if (!t1Covered) pa.pFirstSet.add(r2.learnerId);
      else if (isValidISODate(d1) && isValidISODate(d2) && d2>d1) pa.pFollowSet.add(r2.learnerId);
    }
  }

  // Build pool overview + accumulate EA
  const poolRows=[];
  for (const pa of poolAgg.values()){
    const total=pa.totalSet.size;
    const t2c=pa.t2MonthSet.size;
    const pFirst=pa.pFirstSet.size;
    const pFollow=pa.pFollowSet.size;

    poolRows.push({
      team:pa.team,
      pool:pa.pool,
      ea:pa.ea,
      total_records: total,
      month_connected: t2c,
      month_rate: total? (t2c/total*100).toFixed(1)+"%":"0.0%",
      period_first: pFirst,
      period_followup: pFollow
    });

    const eaKey=k([pa.team, pa.ea]);
    const ea=eaAgg.get(eaKey);
    ea.total += total;
    ea.t2Month += t2c;
    ea.t1Month += pa.t1MonthSet.size;
    ea.pFirst += pFirst;
    ea.pFollow += pFollow;

    const delta = pa.t2MonthSet.size - pa.t1MonthSet.size;
    if (delta !== 0){
      poolDeltaByEA.get(eaKey).set(pa.pool, delta);
    }
  }

  const eaRows=[];
  for (const [eaKey, ea] of eaAgg.entries()){
    const total=ea.total;
    const added = ea.t2Month - ea.t1Month;

    const poolDelta = poolDeltaByEA.get(eaKey) || new Map();
    const parts=[];
    const order=["m2","expiring","duration","period","expired"];
    for (const p of order){
      if (!poolDelta.has(p)) continue;
      const v=poolDelta.get(p);
      if (v>0){
        const label = (p==="m2") ? "M2" : (p==="period" ? "duration" : p);
        parts.push(`${label}-${v}`);
      }
    }
    for (const [p,v] of poolDelta.entries()){
      if (order.includes(p)) continue;
      if (v>0) parts.push(`${p}-${v}`);
    }

    eaRows.push({
      team:ea.team,
      ea:ea.ea,
      total_records: total,
      month_connected: ea.t2Month,
      month_rate: total? (ea.t2Month/total*100).toFixed(1)+"%":"0.0%",
      added_connected: added,
      period_first: ea.pFirst,
      period_followup: ea.pFollow,
      remark: parts.join(", ")
    });
  }

  poolRows.sort((a,b)=>{
    if (a.team!==b.team) return a.team.localeCompare(b.team);
    if (a.pool!==b.pool) return a.pool.localeCompare(b.pool);
    return a.ea.localeCompare(b.ea);
  });
  eaRows.sort((a,b)=> (a.team+a.ea).localeCompare(b.team+b.ea));

  return {eaRows, poolRows};
}

function recencyRank(refDate, monthStart, lastConn){
  if (!isValidISODate(lastConn) || !isValidISODate(refDate)) return 4;
  const days = diffDays(refDate, lastConn);
  if (days >= 21) return 4;
  const inMonth = inRange(lastConn, monthStart, refDate);
  if (!inMonth) return 3;
  if (days > 14) return 2;
  return 1;
}
function recencyLabel(rank){
  if (rank===4) return "21d+ no connect";
  if (rank===3) return "not covered this month";
  if (rank===2) return ">14d in month";
  return "<=14d in month";
}

function buildRecommendations(t2Rows, monthStart, t2End, monthEndFull){
  const refDate = t2End || monthEndFull;
  const byEA=new Map();
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    const pr=poolPriority(r.pool);
    if (pr<=0) continue;
    if (!byEA.has(r.eaName)) byEA.set(r.eaName,[]);
    byEA.get(r.eaName).push(r);
  }

  const out=[];
  const familyKey = (r)=> (r.familyId && r.familyId!=="") ? `F:${r.familyId}` : `L:${r.learnerId}`;

  for (const [ea, rows] of byEA.entries()){
    // Condition 3 filter
    const filtered = rows.filter(r=>{
      const lm = r.lastMonthConsumed;
      const tm = r.thisMonthConsumed;
      if (lm == null || lm < 8) return false;
      if (tm == null || tm <= 0) return false;
      return true;
    });

    const scored = filtered.map(r=>{
      const c1 = recencyRank(refDate, monthStart, r.lastConn);
      const c2 = poolPriority(r.pool);
      const c3 = r.lastMonthConsumed ?? 0;
      const c4 = (r.remaining==null? 999999 : r.remaining); // smaller is better
      return {...r, _c1:c1, _c2:c2, _c3:c3, _c4:c4};
    });

    scored.sort((a,b)=>{
      if (a._c1 !== b._c1) return b._c1 - a._c1;
      if (a._c2 !== b._c2) return b._c2 - a._c2;
      if (a._c3 !== b._c3) return b._c3 - a._c3;
      if (a._c4 !== b._c4) return a._c4 - b._c4;
      return (a.learnerId||"").localeCompare(b.learnerId||"");
    });

    const famMembers=new Map();
    const famBest=new Map();
    for (const r of scored){
      const fk=familyKey(r);
      if (!famMembers.has(fk)) famMembers.set(fk, []);
      famMembers.get(fk).push(r);
      if (!famBest.has(fk)) famBest.set(fk, r);
    }

    const families = Array.from(famBest.entries()).map(([fk,best])=>({fk,best}));
    families.sort((a,b)=>{
      const A=a.best, B=b.best;
      if (A._c1 !== B._c1) return B._c1 - A._c1;
      if (A._c2 !== B._c2) return B._c2 - A._c2;
      if (A._c3 !== B._c3) return B._c3 - A._c3;
      if (A._c4 !== B._c4) return A._c4 - B._c4;
      return (A.learnerId||"").localeCompare(B.learnerId||"");
    });

    let used=0;
    for (const fam of families){
      const members = (famMembers.get(fam.fk) || []).slice().sort((a,b)=>{
        const dp = poolPriority(b.pool) - poolPriority(a.pool);
        if (dp !== 0) return dp;
        return (a.learnerId||"").localeCompare(b.learnerId||"");
      });

      if (used + members.length > 20) continue;

      for (const r of members){
        const reason = `${recencyLabel(r._c1)} | ${r.pool.toUpperCase()} | LM${r.lastMonthConsumed} | Rem${r._c4}`;
        out.push({
          team: r.teamName || "",
          ea,
          learnerId: r.learnerId,
          family_id: r.familyId || "",
          pool: r.pool,
          lastConn: r.lastConn || "",
          lastMonthCons: r.lastMonthConsumed ?? "",
          thisMonthCons: r.thisMonthConsumed ?? "",
          remaining: (r.remaining==null? "" : r.remaining),
          reason
        });
      }
      used += members.length;
      if (used >= 20) break;
    }
  }

  // Keep insertion order; but sort by team then EA for readability without breaking grouping
  out.sort((a,b)=>{
    if (a.team!==b.team) return a.team.localeCompare(b.team);
    if (a.ea!==b.ea) return a.ea.localeCompare(b.ea);
    return 0;
  });

  return out;
}

function chooseExportBase(teamFilter, t2Rows, t2End){
  let team="ALL";
  if (teamFilter && teamFilter.trim()!=="") team = teamFilter.trim();
  else {
    const set=new Set(t2Rows.filter(r=>!POOL_EXCLUDE.has(r.pool)).map(r=>r.teamName).filter(x=>x));
    if (set.size===1) team = Array.from(set)[0];
  }
  team = team.replace(/[\\\/:*?"<>|]/g,"_");
  const datePart = (t2End || "NA");
  return `--${team}-follow-up+${datePart}`;
}

self.onmessage = (ev)=>{
  const data = ev.data || {};
  if (data.type !== "analyze") return;

  try{
    self.postMessage({type:"progress", message:"Parsing workbooks..."});
    const wb1 = readWorkbookFast(data.buf1);
    const wb2 = readWorkbookFast(data.buf2);

    if (!fixedTemplateWorks(wb1) || !fixedTemplateWorks(wb2)){
      self.postMessage({type:"error", message:"Template header mismatch. Required: Learner ID / Assigned SS Staff Name / Last SS Connection Date."});
      return;
    }

    const teamFilter = data.teamFilter || "";
    const t1Rows = parseWorkbook(wb1, teamFilter);
    const t2Rows = parseWorkbook(wb2, teamFilter);

    const monthPick = data.monthPick || "";
    const {start: monthStart, endFull: monthEndFull, label: monthLabel} = monthBoundsFromPicker(monthPick, t2Rows.length?t2Rows:t1Rows);

    const t1End = maxDateInRange(t1Rows, monthStart, monthEndFull);
    const t2End = maxDateInRange(t2Rows, monthStart, monthEndFull);

    self.postMessage({type:"progress", message:"Computing metrics..."});
    const {eaRows, poolRows} = buildMetricsAndRemark(t1Rows, t2Rows, monthStart, monthEndFull, t1End, t2End);
    const recommendations = buildRecommendations(t2Rows, monthStart, t2End, monthEndFull);
    const exportBase = chooseExportBase(teamFilter, t2Rows, t2End);

    self.postMessage({
      type:"result",
      payload:{
        monthLabel,
        t1End,
        t2End,
        exportBase,
        eaOverview: eaRows,
        poolOverview: poolRows,
        recommendations
      }
    });
  }catch(e){
    self.postMessage({type:"error", message: (e && e.message) ? e.message : "Error"});
  }
};
