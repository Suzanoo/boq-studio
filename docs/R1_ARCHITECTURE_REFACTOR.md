# R1 Architecture Refactor

R1 separates BOQ interpretation into five responsibilities without changing the GUI or workbook contract:

1. `core.values` — safe cell conversion and numbering rules
2. `core.extraction` — worksheet cells to typed `ExtractedRow`
3. `core.classification` — row type decision independent of WBS generation
4. `core.wbs` — stateful legacy-compatible WBS projection
5. `core.exporter` — workbook orchestration, Project formulas, Summary and Log

`boq_studio.normalize` remains as a compatibility facade. Existing callers do not need to change.

R1 intentionally preserves unnumbered heading behaviour. The next milestone adds UNKNOWN/confidence and user review before final WBS construction.
