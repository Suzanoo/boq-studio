# MS-5.2 — Item No. Hierarchy Parser

## Goal

Provide one conservative parser for BOQ Item No. hierarchy before building the draft WBS tree.

## Supported forms

- `1`, `1.1`, `1.1.1`
- `1)`, `(1)`, `1.1)`
- `A`, `A.1`, `A.1.1`

## Deliberately rejected

Product/model codes are not treated as hierarchy numbers:

- `F1.1`
- `FDX-301`
- `C1.2`
- `M-01`

Unit or Qty still has priority and makes a row an ITEM regardless of its Item No.

## Output

The parser returns a normalized path, level, parent relation, and sibling relation. MS-5.3 will use these relations to build the draft WBS tree.
