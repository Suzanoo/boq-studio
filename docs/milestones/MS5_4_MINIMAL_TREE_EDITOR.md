# MS-5.4 — Minimal Tree Editor

Adds a minimal, non-persistent draft-tree workspace to each configured sheet.

## Operations

- Indent a node under the previous WBS sibling.
- Outdent a node one level.
- Create a virtual WBS without changing the source workbook.
- Rename WBS nodes.
- Delete virtual WBS nodes while preserving/promoting their children.

WBS depth is capped at two levels below the sheet root (WBS-3 and WBS-4).
Tree persistence and export integration remain MS-5.5/MS-5.6 work.
