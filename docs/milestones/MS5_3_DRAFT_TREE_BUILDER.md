# MS-5.3 — Draft Tree Builder

Builds a first-pass tree for each BOQ sheet from the context-aware row
classification and Item No. hierarchy.

## Rules

- Numbered WBS parents own descendant Item No. rows.
- Text-only WBS headings fall back to Excel indentation and source order.
- ITEM rows attach to their closest numbered WBS parent, otherwise the latest WBS.
- DETAIL rows attach to the preceding ITEM.
- TOTAL and UNKNOWN rows remain visible in the draft; no source row is discarded.
- The tree is an intermediate model only. Editing and persistence are later milestones.
