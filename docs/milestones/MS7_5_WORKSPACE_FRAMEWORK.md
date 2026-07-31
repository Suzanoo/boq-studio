# MS7.5 — Workspace Framework

## Goal

Replace one-off `Toplevel` setup with a shared desktop workspace host while preserving all BOQ, classification, WBS, tree-editing, and export behavior.

## Delivered

- `WorkspaceSpec`: declarative identity, window sizing, modality, and supported UI surfaces.
- `WorkspaceServices`: shared command and layout dependencies.
- `WorkspaceLifecycle`: close guards and deterministic persistence callbacks.
- `WorkspaceWindow`: common Tkinter workspace host.
- Migrated `SheetMappingDialog`, `ClassificationReviewDialog`, and `DraftTreeDialog` to the shared host.
- Draft Tree continues to use the MS7.1 Command Registry, MS7.2 Layout Manager, MS7.3 View State Manager, and MS7.4 desktop UI surfaces.

## Architecture Contract

```text
WorkspaceWindow
├── WorkspaceSpec
├── WorkspaceServices
│   ├── CommandRegistry
│   └── LayoutManager (optional)
├── WorkspaceLifecycle
└── Business-specific workspace widgets
```

Future Normalize Review and Mapping workspaces must use the same host rather than creating independent window infrastructure.

## Non-goals

- No classification behavior changes.
- No Draft Tree hierarchy changes.
- No normalize/export contract changes.
- No new user-facing feature.
