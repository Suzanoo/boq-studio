# MS7.3 — View State Manager

## Goal

Preserve the user's visual working context while the Draft WBS Tree is rebuilt after editing, undo, redo, rename, move, indent, outdent, type conversion, creation, or deletion.

## Delivered

- Added `ViewStateManager` and `ViewStateStore` in `boq_studio.view_state`.
- Preserves expanded and collapsed tree branches.
- Preserves multi-selection and keyboard focus.
- Preserves horizontal and vertical scroll positions.
- Preserves tree column widths.
- Preserves Draft Tree window size and position.
- Stores state separately for each workbook and sheet.
- Ignores deleted or unavailable node IDs safely during restoration.
- Allows an edit operation to override selection with its current target without changing unrelated branch state.

## Refresh flow

```text
Tree edit command
    ↓
Capture current view state
    ↓
Rebuild Treeview rows
    ↓
Restore expanded branches, selection, focus, scroll and columns
    ↓
Continue editing at the same visual position
```

## Persistence

The application stores view state in the operating system's application configuration directory:

- Windows: `%APPDATA%/BOQ Studio/view_state.json`
- macOS: `~/Library/Application Support/BOQ Studio/view_state.json`
- Linux: `$XDG_CONFIG_HOME/boq-studio/view_state.json`

The persistence key includes the resolved workbook path and sheet name, preventing one BOQ sheet from overwriting another sheet's workspace state.

## Acceptance result

- Collapsing Floor and Wall remains unchanged while editing Ceiling.
- Refresh no longer expands the complete tree.
- Selection and focus remain on the edited node.
- Undo and redo retain the current workspace context.
- Window geometry and column widths return when the same workbook sheet is reopened.
- Full automated suite: `126 passed`.

## Manual test checklist

1. Open a Draft WBS Tree containing Floor, Wall, and Ceiling branches.
2. Collapse Floor and Wall.
3. Expand Ceiling and select a Ceiling item.
4. Rename, move, indent, or change the selected item type.
5. Confirm Floor and Wall remain collapsed.
6. Confirm the edited row remains selected and visible.
7. Resize columns and scroll horizontally and vertically.
8. Save or close the Draft Tree and reopen the same workbook sheet.
9. Confirm window geometry, branches, selection, scroll positions, and column widths are restored.
10. Delete a selected row and confirm stale node IDs do not cause an error.
