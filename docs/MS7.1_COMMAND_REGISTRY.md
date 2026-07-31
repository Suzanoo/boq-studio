# MS7.1 — Command Registry

## Status

Complete.

## Goal

Provide one command-routing layer so buttons, menus, context menus, and keyboard shortcuts invoke the same application command instead of duplicating UI logic.

## Architecture

```text
Button ───────────┐
Menu ─────────────┤
Context menu ─────┼──> CommandRegistry ───> Existing application method
Keyboard shortcut ┘
```

The command layer has no Tkinter dependency. Tkinter widgets receive callbacks created by the registry, while all BOQ classification, tree editing, normalization, export, and session logic remains unchanged.

## New modules

```text
src/boq_studio/commands/
├── __init__.py
├── ids.py
└── registry.py
```

### `CommandID`

Stable namespaced identifiers such as:

- `file.open`
- `sheet.configure`
- `tree.delete`
- `workbook.normalize`

### `CommandRegistry`

Responsibilities:

- Register a command handler.
- Prevent accidental duplicate registration.
- Execute commands by stable ID.
- Expose optional enabled-state predicates.
- Produce zero-argument callbacks for Tk buttons and menus.
- Produce event callbacks for keyboard bindings.

## Integrated command sources

### Main application

- File chooser button and File > Open.
- Save Project and application close.
- Find/select/clear sheets.
- Configure Mapping.
- Open Draft WBS Tree.
- Use/Ignore Sheet context-menu actions.
- Normalize workbook.
- About and recent-file clearing.

### Draft WBS Tree

- Editing command-bar buttons.
- Save Draft, Cancel, and Use this tree.
- Double-click, Enter, and F2 rename actions.
- Undo/Redo, Save, Find/Focus, Select All.
- Move, Indent, Outdent, Delete, Escape, and window close.

## Business-logic contract

MS7.1 does not change:

- Row classification.
- WBS tree rules.
- Undo/redo behavior.
- Delete confirmation.
- Normalization output.
- Excel export columns.
- Session persistence.

## Adding a new command

1. Add a stable ID to `CommandID`.
2. Register the existing handler in the owning workspace/application.
3. Connect every UI source using `registry.callback(...)` or `registry.event_callback(...)`.
4. Add routing and regression tests.

Example:

```python
commands.register(CommandID.TREE_DELETE, self._delete)
button = ttk.Button(parent, command=commands.callback(CommandID.TREE_DELETE))
window.bind(
    "<Delete>",
    commands.event_callback(CommandID.TREE_DELETE),
    add="+",
)
```

## Acceptance results

- One registry implementation shared by the main application and Draft Tree workspace.
- Existing business methods remain the command handlers.
- Buttons, menus, context menus, and shortcuts route through command IDs.
- Disabled and unknown-command behavior is covered by unit tests.
- Existing regression suite remains passing.
