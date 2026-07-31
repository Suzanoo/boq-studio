# MS6.4.5 — Draft Tree UI Polish

## Goal

Refactor the Draft WBS Tree footer into a responsive desktop command bar without changing tree-editing behavior.

## Preserved behavior

- Indent / Outdent
- Move Up / Move Down
- WBS ↔ ITEM
- Create WBS
- Edit Description
- Delete Row
- Save Draft / Cancel / Use this tree
- Enter and F2 continue to edit Description
- Existing keyboard shortcuts and Undo / Redo remain unchanged

## Layout

- Wide window: Editing, Status, and Actions groups share one row.
- Compact window: Editing commands wrap into two rows; Status and Actions move below.
- No command is hidden or dropped in either mode.
- Status shows total nodes and WBS / Item / Other counts.
- Footer shows editing state, workbook name, and sheet name.

## Tests

`tests/test_ms6_4_5_draft_tree_ui_polish.py` validates:

- all existing commands remain present and ordered;
- responsive breakpoint behavior;
- wide and compact grid contracts;
- no hidden or duplicate command positions;
- tree summary counts;
- Enter remains reserved for editing rather than approving.

Full suite result: `101 passed`.
