# MS7.6 — Production Polish

## Delivered

- Central semantic appearance profile with light and dark-ready palettes.
- Platform typography policy for Windows, macOS, and Linux.
- High-DPI scaling with safe clamping and Windows DPI-awareness bootstrap.
- Shared ttk styling for the application and every WorkspaceWindow.
- Consistent spacing, button padding, tree row height, headings, and selected rows.
- Tooltip styling sourced from the same appearance profile.
- Accessibility contract: visible keyboard focus, text labels, minimum targets, and no color-only state.

## Compatibility

Business logic, WBS hierarchy, normalization, export, command routing, layout state, and view state are unchanged.

## Dark mode readiness

MS7.6 introduces semantic color tokens and a complete dark palette. Automatic live dark-mode switching is intentionally deferred to user settings so production behavior remains predictable.
