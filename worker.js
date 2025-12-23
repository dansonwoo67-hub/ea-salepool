
importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

const POOL_EXCLUDE = new Set(["m1"]);
const FIXED = {
  learnerId: "Learner ID",
  familyId: "Family ID",
  eaName: "Latest (Current) Assigned SS Staff Name",
  teamName: "Latest (Current) Assigned SS Group Name",
  lastConn: "Last SS Connection Date",
  lastMonthConsumed: "Class Consumption Last Month",
  remaining: "Total Session Card Count"
};

function norm(s){ return (s??"").toString().trim().toLowerCase(); }
function poolNameFromSheet(sheetName){
  const n = norm(sheetName);
  if (n.includes("m2")) return "m2";
  if (n.includes("expiring")) return "expiring";
  if (n.includes("expired")) return "expired";
  if (n === "exp" || n.startsWith("exp ")) return "expired";
  if (n.includes("duration")) return "duration";
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
      const rem = idx.remaining>=0 ? safeNum(r[idx.remaining]) : null;

      rows.push({
        pool,
        learnerId: learnerId.toString().trim(),
        familyId: (familyId??"").toString().trim(),
        eaName: eaName.toString().trim(),
        teamName: (teamName??"").toString().trim(),
        lastConn,
        lastMonthConsumed: lm,
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
  if (pool==="expired") return 2;
  if (pool==="period" || pool==="duration") return 1;
  return 0;
}

function buildOverviews(t1Rows, t2Rows, monthStart, monthEndFull, t1End, t2End){
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

  // define period bounds within month
  const periodStart = t1End ? t1End : addDays(monthStart, -1); // so (periodStart, t2End] starts at monthStart if no t1End
  const periodStartNext = addDays(periodStart, 1);

  for (const [key, r2] of t2Map.entries()){
    const [team,ea,pool] = key.split("||").slice(0,3);
    const poolKey=k([team,pool,ea]);
    const eaKey=k([team,ea]);

    if (!poolAgg.has(poolKey)){
      poolAgg.set(poolKey,{
        team,pool,ea,totalSet:new Set(),
        t1MonthSet:new Set(), t2MonthSet:new Set(),
        t2DaySet:new Set(),
        periodFirstSet:new Set(),
        periodFollowSet:new Set()
      });
    }
    if (!eaAgg.has(eaKey)){
      eaAgg.set(eaKey,{team,ea,total:0,t1Month:0,t2Month:0,t2Day:0,periodFirst:0,periodFollow:0});
    }

    const pa=poolAgg.get(poolKey);
    pa.totalSet.add(r2.learnerId);

    const r1=t1Map.get(key);
    const d1=r1 ? r1.lastConn : null;
    const d2=r2.lastConn;

    const t1Covered = t1End ? inRange(d1, monthStart, t1End) : false;
    const t2Covered = t2End ? inRange(d2, monthStart, t2End) : false;

    if (t2Covered) pa.t2MonthSet.add(r2.learnerId);
    if (t2End && isValidISODate(d2) && d2===t2End) pa.t2DaySet.add(r2.learnerId);
    if (t1Covered) pa.t1MonthSet.add(r2.learnerId);

    // Period increment classification within (t1End, t2End]
    if (t2End && inRange(d2, periodStartNext, t2End) && t2Covered){
      if (!t1Covered){
        pa.periodFirstSet.add(r2.learnerId);
      } else if (isValidISODate(d1) && isValidISODate(d2) && d2>d1){
        pa.periodFollowSet.add(r2.learnerId);
      }
    }
  }

  const poolRows=[];
  for (const pa of poolAgg.values()){
    const total=pa.totalSet.size;
    const t1c=pa.t1MonthSet.size;
    const t2c=pa.t2MonthSet.size;
    const delta=t2c-t1c;
    const t2Day=pa.t2DaySet.size;

    const pFirst=pa.periodFirstSet.size;
    const pFollow=pa.periodFollowSet.size;
    const share = (pFirst+pFollow)>0 ? (pFollow/(pFirst+pFollow)*100).toFixed(1)+"%" : "0.0%";

    poolRows.push({
      team:pa.team,
      pool:pa.pool,
      ea:pa.ea,
      total_records: total,
      t1_month_connected: t1c,
      t1_month_rate: total? (t1c/total*100).toFixed(1)+"%":"0.0%",
      t2_month_connected: t2c,
      t2_month_rate: total? (t2c/total*100).toFixed(1)+"%":"0.0%",
      delta_connected: delta,
      delta_rate: total? (delta/total*100).toFixed(1)+"%":"0.0%",
      t2_latest_day: (t2End||""),
      t2_latest_day_connected: t2Day,
      t2_latest_day_rate: total? (t2Day/total*100).toFixed(1)+"%":"0.0%",
      period_first: pFirst,
      period_followup: pFollow,
      period_followup_share: share
    });

    const eaKey=k([pa.team, pa.ea]);
    const ea=eaAgg.get(eaKey);
    ea.total += total;
    ea.t1Month += t1c;
    ea.t2Month += t2c;
    ea.t2Day += t2Day;
    ea.periodFirst += pFirst;
    ea.periodFollow += pFollow;
  }

  const eaRows=[];
  for (const ea of eaAgg.values()){
    const total=ea.total;
    const delta=ea.t2Month-ea.t1Month;
    const share = (ea.periodFirst+ea.periodFollow)>0 ? (ea.periodFollow/(ea.periodFirst+ea.periodFollow)*100).toFixed(1)+"%" : "0.0%";
    eaRows.push({
      team:ea.team,
      ea:ea.ea,
      total_records: total,
      t1_month_connected: ea.t1Month,
      t1_month_rate: total? (ea.t1Month/total*100).toFixed(1)+"%":"0.0%",
      t2_month_connected: ea.t2Month,
      t2_month_rate: total? (ea.t2Month/total*100).toFixed(1)+"%":"0.0%",
      delta_connected: delta,
      delta_rate: total? (delta/total*100).toFixed(1)+"%":"0.0%",
      t2_latest_day: (t2End||""),
      t2_latest_day_connected: ea.t2Day,
      t2_latest_day_rate: total? (ea.t2Day/total*100).toFixed(1)+"%":"0.0%",
      period_first: ea.periodFirst,
      period_followup: ea.periodFollow,
      period_followup_share: share
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

function buildRecommendations(t2Rows, monthStart, t2End){
  const byEA=new Map();
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    if (!byEA.has(r.eaName)) byEA.set(r.eaName,[]);
    byEA.get(r.eaName).push(r);
  }

  const out=[];
  for (const [ea, rows] of byEA.entries()){
    const candidates = rows.filter(r=>{
      if (!t2End) return true;
      if (r.pool==="m2") return !inRange(r.lastConn, monthStart, t2End);
      if (r.pool==="expiring" || r.pool==="expired") return !inRange(r.lastConn, monthStart, t2End);
      if (r.pool==="period" || r.pool==="duration") return !inRange(r.lastConn, monthStart, t2End);
      return false;
    }).map(r=>{
      let score=0;
      const p=poolPriority(r.pool);
      if (r.pool==="m2"){
        score = 1e9 + p*1e6 + (isValidISODate(r.lastConn)? -Number(r.lastConn.replaceAll("-","")) : 0);
      } else if (r.pool==="expiring" || r.pool==="expired"){
        score = 5e8 + p*1e6 + ((r.lastMonthConsumed??0)*1000);
      } else {
        const rem = (r.remaining==null?999999:r.remaining);
        const remScore = rem<=0?1e6:(1e6/(rem+1));
        score = 2e8 + p*1e6 + remScore + ((r.lastMonthConsumed??0)*500);
      }
      return {...r, _score:score};
    });

    candidates.sort((a,b)=> b._score-a._score);

    const familyKey = (r)=> (r.familyId && r.familyId!=="") ? `F:${r.familyId}` : `L:${r.learnerId}`;
    const famPicked=new Set();
    let famRank=0;

    for (const cand of candidates){
      const fk = familyKey(cand);
      if (famPicked.has(fk)) continue;
      if (famRank>=20) break;

      famPicked.add(fk);
      famRank += 1;

      const members = rows
        .filter(r=>!POOL_EXCLUDE.has(r.pool) && familyKey(r)===fk)
        .sort((a,b)=>{
          const dp=poolPriority(b.pool)-poolPriority(a.pool);
          if (dp!==0) return dp;
          return (a.learnerId||"").localeCompare(b.learnerId||"");
        });

      const team = members[0]?.teamName || cand.teamName || "";
      for (const m of members){
        const reason = (m.learnerId===cand.learnerId && m.pool===cand.pool)
          ? (cand.pool==="m2" ? "M2: not covered in selected month"
             : (cand.pool==="expiring" ? "Expiring: high last-month consumption, not covered in selected month"
             : (cand.pool==="expired" ? "Expired: high last-month consumption, not covered in selected month"
             : "Period/Duration: low remaining, not covered in selected month")))
          : "Same Family ID";

        out.push({
          team,
          ea,
          family_id: m.familyId || "",
          family_rank: famRank,
          learnerId: m.learnerId,
          pool: m.pool,
          output: `${ea} - ${m.learnerId} - ${m.pool}`,
          reason
        });
      }
    }
  }

  out.sort((a,b)=>{
    if (a.team!==b.team) return a.team.localeCompare(b.team);
    if (a.ea!==b.ea) return a.ea.localeCompare(b.ea);
    if (a.family_rank!==b.family_rank) return a.family_rank-b.family_rank;
    const dp=poolPriority(b.pool)-poolPriority(a.pool);
    if (dp!==0) return dp;
    return (a.learnerId||"").localeCompare(b.learnerId||"");
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
    const {eaRows, poolRows} = buildOverviews(t1Rows, t2Rows, monthStart, monthEndFull, t1End, t2End);
    const recommendations = buildRecommendations(t2Rows, monthStart, t2End);
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
