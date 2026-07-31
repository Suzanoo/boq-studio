# MS-6.2 — Persistent Tree Draft

BOQ Studio now treats a workbook plus its mapping and WBS trees as a resumable project.

## Workflow states

- `NEW`
- `MAPPED`
- `TREE DRAFT`
- `TREE READY`
- `EXPORTED`

A saved draft cannot be normalized until the user reopens it and clicks **Use this tree**.

## Tree Editor

- **Save Draft** stores the current tree without approving it.
- **Use this tree** stores and approves the tree.
- Closing with unsaved changes offers Save Draft, Discard, or Continue Editing.

## Resume behavior

When the user selects a BOQ workbook, BOQ Studio looks for its hidden project session. If the workbook fingerprint matches, the user can resume mapping, draft trees, ready states, sheet selections, and export states. If the workbook changed, the old session is not restored.

## Storage

Sessions are stored under the operating system application-data folder and written atomically as JSON.
