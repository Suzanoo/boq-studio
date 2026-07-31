# MS-6.4.1 — Delete Key and Cost Column Hotfix

## Fixes

- Bind Delete and keypad Delete at the Draft Tree window level so the shortcut works even after focus briefly moves to another child widget.
- Keep mouse selection focused on the tree.
- Preserve Delete confirmation and Ctrl+Z undo behavior.
- Restore Material and Labor to the public workbook export contract.

## Export contract

1. WBS No.
2. WBS-1
3. WBS-2
4. WBS-3
5. WBS-4
6. Description
7. Unit
8. Qty
9. Material
10. Labor
11. Amount

The normalized sheets and Project sheet use the same 11-column order. Project values remain formula-linked to their source normalized sheets. Summary calculations use Amount in column K.

## Verification

- Delete shortcut dispatch test
- Material, Labor, and Amount export test
- Project formula-link regression test
- Excel outline and print-area regression test
- Full test suite: 75 passed
