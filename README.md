# EA Coverage Dashboard (v10)

Enhancement:
- Adds BOTH metrics within (T1 end, T2 end] for the selected month:
  - period_first: newly covered in the month during the period (not covered by T1 end, covered by T2 end)
  - period_followup: updated follow-up connects during the period (covered by T1 end, and T2 date is later than T1 date)
  - period_followup_share: followup / (first + followup)

Other rules unchanged:
- Month comparison is computed on the T2 roster (denominator = total_records from T2).
- Latest-day coverage uses T2 end date within selected month.
- M1 excluded.
