# EA Coverage Dashboard (v14)

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
