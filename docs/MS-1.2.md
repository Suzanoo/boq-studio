# MS-1.2 — OKD row filtering and shallow WBS

- Export only cost rows where `Amount > 0`.
- Use the source sheet name as `WBS-1`.
- Use the latest detected BOQ heading as `WBS-2`.
- Keep `WBS-3` and `WBS-4` blank for this milestone.
- Apply the same rules to Normalized Preview and Excel export.
- Preserve `Source Sheet` and `Source Row` for traceability.

The program continues to parse heading rows internally, but does not export those
rows because they have no amount.
