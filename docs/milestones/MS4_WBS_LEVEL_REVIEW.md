# MS-4 — WBS Level Review

MS-4 extends the classification workspace so a reviewed heading can be assigned
explicitly to WBS-2, WBS-3 or WBS-4.

## Data contract

`SheetConfig.wbs_level_overrides` stores `{source_row: level}` where level is 2,
3 or 4. Mapping templates use version 3 and remain backward compatible with
older templates that do not contain this field.

## Export behavior

- WBS-2 replaces the sheet-name default and clears WBS-3/WBS-4.
- WBS-3 clears the active WBS-4.
- WBS-4 changes only the lowest heading.
- ITEM rows inherit the current reviewed hierarchy.
