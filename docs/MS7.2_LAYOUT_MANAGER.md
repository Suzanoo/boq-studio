# MS7.2 — Layout Manager

## Status

Complete.

## Goal

Introduce one reusable layout controller for BOQ Studio workspaces and make the Draft WBS Tree sidebar collapsible without changing BOQ classification, tree-editing, or export logic.

## Architecture

```text
Workspace UI
├── Main content
├── Sidebar
├── Command bar
└── Status bar
        │
        ▼
LayoutManager
        │
        ├── visibility state
        ├── sidebar width
        └── LayoutStateStore
                │
                ▼
        OS user configuration/layout.json
```

The core layout classes do not import Tkinter. Tk widgets are connected through small `GridRegion` and `PackRegion` adapters.

## New module

```text
src/boq_studio/layout.py
```

### `LayoutState`

Stores workspace preferences:

- Sidebar visibility.
- Command-bar visibility.
- Status-bar visibility.
- Toolbar visibility for future shared workspaces.
- Sidebar width with safe minimum and maximum limits.

### `LayoutStateStore`

- Loads and saves JSON state by workspace ID.
- Keeps states for different workspaces separate.
- Falls back safely when the file is missing or invalid.
- Preserves unrelated workspace entries.

### `LayoutManager`

- Registers named UI regions.
- Applies persisted state immediately.
- Shows, hides, or toggles regions.
- Persists every layout change.
- Rejects unknown region names instead of silently failing.

## Draft WBS Tree integration

The Draft WBS Tree now has:

- A right-side **Node Details** panel.
- A close control inside the sidebar.
- A compact restore control when the sidebar is collapsed.
- Selected-node description, type, WBS level, amount, and source row.
- Persistent sidebar state between application launches.
- Layout commands routed through the MS7.1 Command Registry.

The command bar and status bar are also registered with the same Layout Manager, ready for the standard View menu in MS7.4.

## Regression repair

The Draft WBS Tree now explicitly creates and registers its `CommandRegistry` before building widgets. This preserves the MS7.1 routing contract during real GUI use.

## Persistence locations

- Windows: `%APPDATA%/BOQ Studio/layout.json`
- macOS: `~/Library/Application Support/BOQ Studio/layout.json`
- Linux: `$XDG_CONFIG_HOME/boq-studio/layout.json`

## Business-logic contract

MS7.2 does not change:

- Row classification.
- Draft-tree construction.
- Tree editing and Undo/Redo.
- WBS numbering.
- Material, Labor, or Amount export.
- Workbook normalization.
- Project-session persistence.

## Acceptance results

- Sidebar collapses and restores without rebuilding the Draft Tree.
- Layout choice is remembered across application launches.
- Window resizing leaves the main workspace available.
- All layout actions use shared Command IDs.
- Layout logic is testable without opening Tkinter.
- Existing regression suite remains passing.
