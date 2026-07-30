# MS-3 — Classification Review Workspace

## Goal
Let users confirm ambiguous BOQ rows before WBS generation without changing the source workbook.

## Delivered
- Review workspace launched from the selected sheet.
- Displays source row, description, automatic type, final type, confidence, reason, Unit, Qty and Amount.
- Filters: Review needed, All rows, Overrides only and UNKNOWN only.
- Batch actions: Mark WBS, Mark ITEM, Ignore and Reset Auto.
- Per-sheet overrides persist in mapping templates.
- Export and normalized preview honor reviewed classifications.
- User-confirmed text-only WBS rows are accepted as WBS-3 until WBS level review is implemented.

## Compatibility
Existing mapping templates remain loadable. Templates saved by MS-3 use schema version 2.
