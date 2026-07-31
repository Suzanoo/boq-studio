# MS7.4 — Standard Desktop UI

## Scope

Draft WBS Tree now exposes one command architecture through standard desktop surfaces:

- Menu Bar: File, Edit, Tree, View
- Compact Toolbar for frequent commands
- Context Menu for selected tree nodes
- Lazy Tooltips with native Windows/macOS shortcut labels
- Existing Status Bar and Command Bar remain managed by LayoutManager

## Safety

No BOQ, classification, hierarchy, normalization, or export business logic was changed. Menu, toolbar, context menu, keyboard shortcuts, and existing buttons all route through `CommandRegistry`.

## View commands

- Expand Selected / Collapse Selected
- Expand All / Collapse All

These change only Treeview presentation state and do not modify the DraftTree model or Undo stack.

## Tests

`tests/test_ms7_4_standard_desktop_ui.py` validates command coverage, duplicate prevention, toolbar/context-menu contracts, paired tree view commands, and native shortcut/tooltip formatting.
