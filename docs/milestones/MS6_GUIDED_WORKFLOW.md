# MS-6 — Guided Workflow

BOQ Studio now guides users through four required stages:

1. Select BOQ File
2. Configure Mapping
3. Approve Draft WBS Tree
4. Normalize to Desktop

Each selected sheet has one state:

- `NEW` — mapping is missing
- `MAPPED` — mapping exists, but the Draft WBS Tree has not been approved
- `TREE READY` — the user clicked **Use this tree**
- `EXPORTED` — the sheet was included in a completed export

The Normalize button remains disabled until every selected sheet is `TREE READY` or `EXPORTED`.
Reconfiguring a sheet invalidates its previous Draft Tree so stale structure cannot be exported.
