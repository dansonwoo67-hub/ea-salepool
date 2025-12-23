# EA Coverage Dashboard (v16)

Fixes per your feedback:
1) EA overview:
   - added_connected = T2 latest day (t2End) connected count (exclude M1).
   - remark = pool breakdown of added_connected on t2End (sum equals added_connected).
   - Renamed:
     - month_first = (t1End, t2End] first-time connects in the selected month
     - month_follow_up = (t1End, t2End] follow-up connects (date updated later than T1)
2) By pool:
   - Only T2 month metrics + month_first/month_follow_up.
3) Recommended IDs:
   - Main logic: Pool priority first, then NOT covered this month, then older last connect, then higher last-month consumption, then lower remaining.
   - Hard filters: last-month consumption >= 8 AND this-month consumption > 0; M1 excluded.
   - Do not recommend "covered this month" unless pool coverage is already very high:
     - include covered (>14d) only if pool month coverage >= 50%
     - include covered (<=14d) only if pool month coverage >= 65%
   - Keep Family ID together; max 20 IDs per EA (hard cap; skip a family if it would exceed 20).


v14 updates:
- Recommended IDs: exclude records connected within the last 7 days (even if pool coverage is high).
- By pool table: added `added_connected` (T2 latest-day connects per pool).


v15 updates:
- Recommendations: "not connected this month" always prioritized.
- Family rule: if a family contains ANY record connected this month, the whole family is excluded from the "uncovered" bucket; mixed families (some covered, some not) are fully excluded (no recommendations for that family).
- Follow-up gating uses OVERALL EA month coverage (all pools excluding M1), not per-pool coverage.


v16 updates:
- Recommendations: family-level exclusion if ANY member connected within last 14 days (relative to T2 end date).
- Recommendations: still exclude mixed families (some connected this month, some not).
- By pool: `added_connected` now counts connections on the pool’s own latest day within the selected month window.
- Cache bust: worker is loaded as worker.js?v=16 to avoid stale caching.
