# MS-2 — Row Classification

MS-2 separates row interpretation from WBS construction and exposes uncertainty for review.

## Added

- `ClassificationResult(row_type, confidence, reason)`
- `UNKNOWN` for ambiguous text-only rows
- `ClassificationReviewService` and `ReviewRow`
- Structural total detection independent of Thai/English keywords
- Regression tests for totals, items containing `รวม`, WBS headings, and unknown rows

## Compatibility

The existing exporter continues to run. Until the review UI is implemented, `UNKNOWN` rows are passed to the legacy WBS heading path so current outputs do not silently change.
