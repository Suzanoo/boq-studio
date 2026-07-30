# MS-5.1 — Context-aware classifier

The classifier now evaluates a sequence of BOQ rows rather than assuming that
an Amount-only row is always a total.

Key rules:

- Unit or Qty present: ITEM.
- A numbered row followed by numbered descendants: WBS.
- A heading immediately before item rows: WBS candidate.
- An Amount-only row after items and before the next group/end of sheet: TOTAL.
- Dash/bullet text without quantities: DETAIL.
- Ambiguous Amount-only rows remain UNKNOWN for the future tree review.

This milestone changes detection only. The Draft WBS Tree UI is implemented in
the next milestone.
