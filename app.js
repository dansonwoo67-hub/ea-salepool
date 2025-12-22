/* ==========
  EA Coverage Dashboard (Static / GitHub Pages)
  - Load 2 Excel files (T1 old, T2 new)
  - Manual date selection for each
  - Compute coverage by Team/EA/Pool and delta
  - Identify first-round vs follow-up changes within interval
  - Recommend up to 20 follow-up IDs per EA based on priority rules
========== */

const $ = (id) => document.getElementById(id);

let wb1 = null, wb2 = null;
let data1 = null, data2 = null;
let detectedColumns = null;
let mapping = null;

const POOL_EXCLUDE = new Set(["m1"]);
const DEFAULT_POOLS = ["m2", "expiring", "expired", "duration", "period", "exp"]; // "exp" optional naming

// Column canonical keys
const COLS = {
  learnerId: "learnerId",
  eaName: "eaName",
  teamName: "teamName",
  lastConn: "lastConn",
  lastFollow: "lastFollow",
  remaining: "remaining",
  lastMonthConsumed: "lastMonthConsumed"
};

// Heuristic alias lists (case-insensitive contains match)
const ALIASES =
