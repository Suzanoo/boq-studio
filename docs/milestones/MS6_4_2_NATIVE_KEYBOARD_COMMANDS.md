# MS6.4.2 — Native Keyboard Commands (Priority A)

## Scope

BOQ Studio now routes the essential desktop commands through OS-appropriate shortcuts.

| Command | Windows / Linux | macOS |
|---|---|---|
| Open workbook | Ctrl+O | Command+O |
| Save project/draft | Ctrl+S | Command+S |
| Find / focus current table | Ctrl+F | Command+F |
| Select all sheets | Ctrl+A | Command+A |
| Undo tree edit | Ctrl+Z | Command+Z |
| Redo tree edit | Ctrl+Y or Ctrl+Shift+Z | Command+Shift+Z |
| Delete selected tree rows | Delete / keypad Delete | Delete |
| Rename selected tree row | Enter / F2 / double-click | Return / F2 / double-click |
| Cancel or close draft editor | Escape | Escape |

## Safety contract

- Delete acts on selected Draft Tree rows only when focus is not inside a text-entry widget.
- Entry, Text, Spinbox, and Combobox widgets retain native Delete, Copy, Cut, Paste, and Select All behavior.
- Toolbar actions and keyboard shortcuts call the same editor operations.
- Undo/Redo covers destructive and structural Draft Tree edits.
- Dialog shortcuts are scoped to the dialog and are removed when it closes.

## Automated verification

`tests/test_ms6_4_2_native_shortcuts.py` verifies:

1. Windows shortcut mapping.
2. macOS shortcut mapping.
3. Delete safety for text-input widgets.
4. Delete eligibility for Treeview/window widgets.

Full regression result: `81 passed`.
