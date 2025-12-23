# EA Coverage Dashboard (v11)

Implements your optimization requirements:
1) EA overview: only T2 month metrics + added_connected vs T1, keep period_first + period_followup, add remark like "M2-3, duration-4".
2) By pool: only T2 metrics + period_first + period_followup (no T1 columns shown).
3) Recommended IDs (max 20 per EA):
   - Priority (descending): recency bucket > pool priority > last-month consumption > remaining sessions.
   - Filters: last month consumption >= 8 AND this month consumption > 0; M1 excluded.
   - Keep same Family ID together; skip a family if it would exceed 20 IDs.
