# MS6.4.4 — Native OS Integration (Priority C)

## Scope delivered

- Native File, Edit, and Help menus.
- macOS application menu with About and Quit integration.
- Windows Exit and macOS Close/Quit behavior through one safe close command.
- Open Recent workbook list persisted in the OS-appropriate user config folder.
- Right-click sheet context menu; Control-click is also supported on macOS.
- Existing standard Open dialog retained.
- Project session is saved before a normal application close.
- A close warning is shown while normalization/export is running.

## Behavior intentionally preserved

- WBS classification and hierarchy algorithms.
- Draft Tree edit commands and Undo/Redo stack.
- Excel export column contract and outline behavior.
- Priority A and Priority B keyboard shortcuts.

## Deferred to Production planning

Drag-and-drop workbook opening is intentionally excluded. Tkinter requires an
additional platform integration package for reliable Windows/macOS file drops.
It should be evaluated together with packaging, signing, and installer design.

## Automated verification

`tests/test_ms6_4_4_native_os_integration.py` covers platform selection,
configuration paths, recent-file filtering/deduplication/limits, persistence,
and most-recent ordering.
