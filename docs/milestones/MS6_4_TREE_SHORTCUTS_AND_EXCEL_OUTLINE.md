# MS-6.4 — Tree Shortcuts and Professional Excel Outline

## Tree editor

- `Ctrl+Z`: undo the latest tree edit
- `Ctrl+Y` or `Ctrl+Shift+Z`: redo
- `Delete`: delete the selected row(s), using the same confirmation as the button
- `Enter`, `F2`, or double-click: edit the selected description
- Undo history covers indent, outdent, move, WBS/ITEM toggle, create WBS, edit description, and delete.
- Up to 100 undo snapshots are retained while the dialog is open.

## Excel export

- Adds native Excel Outline / Row Grouping based on project-wide `WBS No.` depth.
- The sheet root remains outline level 0; descendants receive levels 1–7.
- Sets `summaryBelow = false`, so parent WBS rows appear above their grouped children.
- Copies outline levels to the formula-linked `Project` sheet.
- WBS rows are bold with hierarchical fills; ITEM rows remain unfilled.
- Freeze header, AutoFilter, fixed widths, number formats, landscape print, repeat header row, fit-to-width, print area, and compact margins.
