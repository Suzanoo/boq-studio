# MS6.4.3 — Tree Editing Shortcuts

## Scope

Add keyboard access to the existing Draft Tree commands without changing the
WBS hierarchy engine or replacing native Treeview navigation.

| Command | Windows | macOS |
|---|---|---|
| Move selected row(s) up | Alt+Up | Option+Up |
| Move selected row(s) down | Alt+Down | Option+Down |
| Indent selected row | Alt+Right | Option+Right |
| Outdent selected row | Alt+Left | Option+Left |
| Select all tree rows | Ctrl+A | Command+A |
| Expand / collapse | Right / Left | Right / Left (native Treeview) |

## Safety contract

- Keyboard commands call the same methods as the existing UI buttons.
- Multi-row move behavior and ordering remain in `DraftTreeEditor` unchanged.
- Plain arrow keys remain native Treeview behavior.
- Tree-changing shortcuts do nothing while a text-entry widget has focus.
- Move, indent, and outdent continue to participate in the existing Undo/Redo stack.
- No duplicate-row command was added because no established production command
  existed to reuse safely in this milestone.
