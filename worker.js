
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
  if (n === "exp" || n === "exp " || n.startsWith("exp ")) return "expired"; // EXP sheet = expired (distinct from expiring)
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
  // Treat 0 / "0" placeholders as empty (avoid 1970-01-01 artifacts)
  if (v === 0) return null;
  if (typeof v === "string"){
    const t = v.trim();
    if (t === "" || t === "0" || t === "0.0" || t === "-") return null;
  }

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
  if (s === "0" || s === "0.0" || s === "-") return null;

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


// Fast dense-sheet cell read (dense:true returns ws as row arrays of cell objects)
function getDenseVal(ws, r, c){
  const row = ws ? ws[r] : null;
  if (!row) return null;
  const cell = row[c];
  if (cell == null) return null;
  if (typeof cell === "object" && "v" in cell) return cell.v;
  return cell;
}
function denseRowToHeader(ws, range){
  const hr = range.s.r;
  const header = [];
  for (let c=range.s.c; c<=range.e.c; c++){
    const v = getDenseVal(ws, hr, c);
    header.push(v == null ? "" : v.toString());
  }
  return {header, hr};
}

function headersContain(wb, col){
  for (const sn of wb.SheetNames){
    const ws = wb.Sheets[sn];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const hr = range.s.r;
    for (let c=range.s.c; c<=range.e.c; c++){
      const v = getDenseVal(ws, hr, c);
      if (v == null) continue;
      if (v.toString() === col) return true;
    }
  }
  return false;
}
function fixedTemplateWorks(wb){
  return headersContain(wb, FIXED.learnerId) && headersContain(wb, FIXED.eaName) && headersContain(wb, FIXED.lastConn);
}
function fixedTemplateWorks(wb){
  return headersContain(wb, FIXED.learnerId) && headersContain(wb, FIXED.eaName) && headersContain(wb, FIXED.lastConn);
}

function parseWorkbook(wb, teamFilter){
  const rows=[];
  for (const sn of wb.SheetNames){
    const pool=poolNameFromSheet(sn);
    const ws=wb.Sheets[sn];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    if (range.e.r - range.s.r < 1) continue;

    const {header, hr} = denseRowToHeader(ws, range);

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

    const baseC = range.s.c;

    for (let r=hr+1; r<=range.e.r; r++){
      const learnerRaw = getDenseVal(ws, r, baseC + idx.learnerId);
      const eaRaw = getDenseVal(ws, r, baseC + idx.eaName);
      if (learnerRaw == null || learnerRaw === "" || eaRaw == null || eaRaw === "") continue;

      const teamRaw = idx.teamName>=0 ? getDenseVal(ws, r, baseC + idx.teamName) : "";
      const teamName = (teamRaw ?? "").toString().trim();
      if (teamFilter && teamName && teamName !== teamFilter.trim()) continue;

      const lastConnRaw = idx.lastConn>=0 ? getDenseVal(ws, r, baseC + idx.lastConn) : null;
      const lastConn = excelDateToISO(lastConnRaw);

      const familyRaw = idx.familyId>=0 ? getDenseVal(ws, r, baseC + idx.familyId) : "";
      const lmRaw = idx.lastMonthConsumed>=0 ? getDenseVal(ws, r, baseC + idx.lastMonthConsumed) : null;
      const tmRaw = idx.thisMonthConsumed>=0 ? getDenseVal(ws, r, baseC + idx.thisMonthConsumed) : null;
      const remRaw = idx.remaining>=0 ? getDenseVal(ws, r, baseC + idx.remaining) : null;

      rows.push({
        pool,
        learnerId: learnerRaw.toString().trim(),
        familyId: (familyRaw??"").toString().trim(),
        eaName: eaRaw.toString().trim(),
        teamName,
        lastConn,
        lastMonthConsumed: safeNum(lmRaw),
        thisMonthConsumed: safeNum(tmRaw),
        remaining: safeNum(remRaw)
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
  const latestDayPoolByEA=new Map();
  const latestDayCountByEA=new Map();

  const periodStart = t1End ? t1End : addDays(monthStart, -1);
  const periodStartNext = addDays(periodStart, 1);

  for (const [key, r2] of t2Map.entries()){
    const [team,ea,pool] = key.split("||").slice(0,3);
    const poolKey=k([team,pool,ea]);
    const eaKey=k([team,ea]);

    if (!poolAgg.has(poolKey)){
      poolAgg.set(poolKey,{team,pool,ea,totalSet:new Set(), t2MonthSet:new Set(), monthFirstSet:new Set(), monthFollowSet:new Set(), latestDaySet:new Set(), latestDate:null});
    }
    if (!eaAgg.has(eaKey)){
      eaAgg.set(eaKey,{team,ea,total:0,t2Month:0,monthFirst:0,monthFollow:0});
    }
    if (!latestDayPoolByEA.has(eaKey)) latestDayPoolByEA.set(eaKey, new Map());
    if (!latestDayCountByEA.has(eaKey)) latestDayCountByEA.set(eaKey, new Set());

    const pa=poolAgg.get(poolKey);
    pa.totalSet.add(r2.learnerId);

    const r1=t1Map.get(key);
    const d1=r1 ? r1.lastConn : null;
    const d2=r2.lastConn;

    const t1Covered = t1End ? inRange(d1, monthStart, t1End) : false;
    const t2Covered = t2End ? inRange(d2, monthStart, t2End) : false;

    if (t2Covered) pa.t2MonthSet.add(r2.learnerId);
    // Pool-level latest day in the selected month window (for added_connected in By pool)
    if (t2Covered && isValidISODate(d2)){
      if (!pa.latestDate || d2 > pa.latestDate){
        pa.latestDate = d2;
        pa.latestDaySet = new Set([r2.learnerId]);
      } else if (d2 === pa.latestDate){
        pa.latestDaySet.add(r2.learnerId);
      }
    }

    if (t2End && inRange(d2, periodStartNext, t2End) && t2Covered){
      if (!t1Covered) pa.monthFirstSet.add(r2.learnerId);
      else if (isValidISODate(d1) && isValidISODate(d2) && d2>d1) pa.monthFollowSet.add(r2.learnerId);
    }

    if (t2End && isValidISODate(d2) && d2===t2End){
      latestDayCountByEA.get(eaKey).add(r2.learnerId);

      const m = latestDayPoolByEA.get(eaKey);
      const existing = m.get(r2.learnerId);
      if (!existing || poolPriority(pool) > poolPriority(existing)){
        m.set(r2.learnerId, pool);
      }
    }
  }

  const poolRows=[];
  for (const pa of poolAgg.values()){
    const total=pa.totalSet.size;
    const t2c=pa.t2MonthSet.size;
    const mf=pa.monthFirstSet.size;
    const mfu=pa.monthFollowSet.size;

    poolRows.push({
      team:pa.team,
      pool:pa.pool,
      ea:pa.ea,
      total_records: total,
      month_connected: t2c,
      month_rate: total? (t2c/total*100).toFixed(1)+"%":"0.0%",
      added_connected: (pa.latestDaySet ? pa.latestDaySet.size : 0),
      month_first: mf,
      month_follow_up: mfu
    });

    const eaKey=k([pa.team, pa.ea]);
    const ea=eaAgg.get(eaKey);
    ea.total += total;
    ea.t2Month += t2c;
    ea.monthFirst += mf;
    ea.monthFollow += mfu;
  }

  const eaRows=[];
  for (const [eaKey, ea] of eaAgg.entries()){
    const total=ea.total;
    const addedSet = latestDayCountByEA.get(eaKey) || new Set();
    const added = addedSet.size;

    const learnerToPool = latestDayPoolByEA.get(eaKey) || new Map();
    const counts=new Map();
    for (const lid of addedSet){
      const p = learnerToPool.get(lid) || "unknown";
      const label = (p==="m2") ? "M2" : (p==="period" ? "duration" : p);
      counts.set(label, (counts.get(label)||0)+1);
    }
    const order=["M2","expiring","duration","expired","unknown"];
    const parts=[];
    for (const p of order){
      if (counts.has(p)) parts.push(`${p}-${counts.get(p)}`);
    }
    for (const [p,c] of counts.entries()){
      if (!order.includes(p)) parts.push(`${p}-${c}`);
    }

    eaRows.push({
      team:ea.team,
      ea:ea.ea,
      total_records: total,
      month_connected: ea.t2Month,
      month_rate: total? (ea.t2Month/total*100).toFixed(1)+"%":"0.0%",
      added_connected: added,
      month_first: ea.monthFirst,
      month_follow_up: ea.monthFollow,
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

function monthCovered(refDate, monthStart, lastConn){
  return inRange(lastConn, monthStart, refDate);
}

function buildCoverageMapByEAandPool(t2Rows, monthStart, refDate){
  const total=new Map();
  const covered=new Map();
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    const pr=poolPriority(r.pool);
    if (pr<=0) continue;
    const key=k([r.teamName, r.eaName, r.pool]);
    if (!total.has(key)) total.set(key, new Set());
    total.get(key).add(r.learnerId);
    if (monthCovered(refDate, monthStart, r.lastConn)){
      if (!covered.has(key)) covered.set(key, new Set());
      covered.get(key).add(r.learnerId);
    }
  }
  const rate=new Map();
  for (const [key, set] of total.entries()){
    const cov = covered.get(key)?.size || 0;
    const tot = set.size || 0;
    rate.set(key, tot ? (cov/tot) : 0);
  }
  return {rate};
}

function buildRecommendations(t2Rows, monthStart, t2End, monthEndFull){
  const refDate = t2End || monthEndFull;

  // Overall EA month coverage across ALL pools (excluding M1). Used for follow-up gating.
  // key: team||ea
  const totalEA = new Map();
  const coveredEA = new Map();
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    if (poolPriority(r.pool) <= 0) continue;
    const eaKey = k([r.teamName, r.eaName]);
    if (!totalEA.has(eaKey)) totalEA.set(eaKey, new Set());
    totalEA.get(eaKey).add(r.learnerId);
    if (monthCovered(refDate, monthStart, r.lastConn)){
      if (!coveredEA.has(eaKey)) coveredEA.set(eaKey, new Set());
      coveredEA.get(eaKey).add(r.learnerId);
    }
  }
  const eaCovRate = new Map();
  for (const [eaKey, set] of totalEA.entries()){
    const tot = set.size || 0;
    const cov = coveredEA.get(eaKey)?.size || 0;
    eaCovRate.set(eaKey, tot ? (cov/tot) : 0);
  }

  const TH_ALLOW_FOLLOWUP = 0.50; // allow follow-up (>14d) only if overall EA coverage >= 50%

  // Group rows by (team, ea)
  const byEA = new Map(); // team||ea -> rows
  for (const r of t2Rows){
    if (POOL_EXCLUDE.has(r.pool)) continue;
    if (poolPriority(r.pool) <= 0) continue;
    const eaKey = k([r.teamName, r.eaName]);
    if (!byEA.has(eaKey)) byEA.set(eaKey, []);
    byEA.get(eaKey).push(r);
  }

  const out = [];
  const familyKey = (r)=> (r.familyId && r.familyId!=="") ? `F:${r.familyId}` : `L:${r.learnerId}`;

  for (const [eaKey, rows] of byEA.entries()){
    const [team, ea] = eaKey.split("||");

    // Hard filter: last month >=8 and this month >0 (otherwise not recommended)
    const filtered = rows.filter(r=>{
      const lm = r.lastMonthConsumed;
      const tm = r.thisMonthConsumed;
      if (lm == null || lm < 8) return false;
      if (tm == null || tm <= 0) return false;
      return true;
    });

    // Build family groups
    const famMembers = new Map();
    for (const r of filtered){
      const fk = familyKey(r);
      if (!famMembers.has(fk)) famMembers.set(fk, []);
      famMembers.get(fk).push(r);
    }

    // Build candidate family info
    const families = [];
    for (const [fk, members] of famMembers.entries()){
      // Sort members in family for stable output (M2 first)
      members.sort((a,b)=>{
        const dp = poolPriority(b.pool) - poolPriority(a.pool);
        if (dp !== 0) return dp;
        return (a.learnerId||"").localeCompare(b.learnerId||"");
      });

      // Determine month-covered flags
      const coveredFlags = members.map(r => monthCovered(refDate, monthStart, r.lastConn));
      const anyCovered = coveredFlags.some(Boolean);
      const anyUncovered = coveredFlags.some(x=>!x);

      // Mixed family exclusion: if one is month-covered, all IDs in family are excluded
      if (anyCovered && anyUncovered) continue;

      // Compute family most recent connect date (any pool)
      let mostRecent = null;
      for (const r of members){
        const d = r.lastConn;
        if (isValidISODate(d) && (!mostRecent || d > mostRecent)) mostRecent = d;
      }
      const daysMostRecent = mostRecent ? diffDays(refDate, mostRecent) : 9999;

      // Rule: if any member connected within last 14 days, exclude the whole family
      if (mostRecent && daysMostRecent <= 14) continue;

      // Bucket logic:
      // bucket 0: not covered this month (always eligible; highest priority)
      // bucket 1: covered this month but >14 days since last connect (eligible only if overall EA coverage >= 50%)
      let bucket = 0;
      if (anyCovered) bucket = 1;

      const covRate = eaCovRate.get(eaKey) || 0;
      if (bucket === 1 && covRate < TH_ALLOW_FOLLOWUP) continue;

      // Best member for scoring: higher pool priority, higher LM, lower remaining, older lastConn
      let best = members[0];
      for (const r of members){
        const ap = poolPriority(r.pool), bp = poolPriority(best.pool);
        if (ap !== bp){ if (ap > bp) best = r; continue; }
        const alm = r.lastMonthConsumed ?? 0, blm = best.lastMonthConsumed ?? 0;
        if (alm !== blm){ if (alm > blm) best = r; continue; }
        const arem = (r.remaining==null? 999999 : r.remaining);
        const brem = (best.remaining==null? 999999 : best.remaining);
        if (arem !== brem){ if (arem < brem) best = r; continue; }
        const ad = isValidISODate(r.lastConn) ? r.lastConn : "0000-00-00";
        const bd = isValidISODate(best.lastConn) ? best.lastConn : "0000-00-00";
        if (ad !== bd){ if (ad < bd) best = r; continue; } // older date preferred
      }

      // Recency sub-score (Condition 1): >21d no connect > otherwise
      const days = mostRecent ? diffDays(refDate, mostRecent) : 9999;
      const recencyTier = (days > 21) ? 0 : 1;

      const reason = bucket===0
        ? `Not connected this month; ${days>21?'>21d':'<=21d'} no connect`
        : `Follow-up needed: last connect >14d`;

      families.push({ fk, members, best, bucket, recencyTier, days, covRate, reason });
    }

    // Sort families: bucket0 first, then recencyTier, then pool priority, then higher LM, then lower remaining, then older connect
    families.sort((A,B)=>{
      if (A.bucket !== B.bucket) return A.bucket - B.bucket;
      if (A.recencyTier !== B.recencyTier) return A.recencyTier - B.recencyTier;
      const ap = poolPriority(B.best.pool) - poolPriority(A.best.pool);
      if (ap !== 0) return ap;
      const lm = (B.best.lastMonthConsumed??0) - (A.best.lastMonthConsumed??0);
      if (lm !== 0) return lm;
      const ar = (A.best.remaining==null? 999999 : A.best.remaining);
      const br = (B.best.remaining==null? 999999 : B.best.remaining);
      if (ar !== br) return ar - br;
      const ad = isValidISODate(A.best.lastConn) ? A.best.lastConn : "0000-00-00";
      const bd = isValidISODate(B.best.lastConn) ? B.best.lastConn : "0000-00-00";
      if (ad !== bd) return ad.localeCompare(bd); // older first
      return (A.best.learnerId||"").localeCompare(B.best.learnerId||"");
    });

    // Emit up to 20 learner IDs per EA (family kept together; if a family would exceed cap, skip it)
    let used = 0;
    for (const f of families){
      const count = f.members.length;
      if (used + count > 20) continue;

      for (const r of f.members){
        out.push({
          team,
          ea,
          learnerId: r.learnerId,
          family_id: r.familyId || "",
          pool: r.pool,
          lastConn: r.lastConn || "",
          lastMonthCons: r.lastMonthConsumed ?? "",
          thisMonthCons: r.thisMonthConsumed ?? "",
          remaining: (r.remaining==null? "" : r.remaining),
          reason: f.reason
        });
      }
      used += count;
      if (used >= 20) break;
    }
  }

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
