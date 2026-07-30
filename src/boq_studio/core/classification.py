from __future__ import annotations

from enum import StrEnum

from .extraction import ExtractedRow
from .values import is_blank_value, numbering_level

NOTE_WORDS = ("หมายเหตุ", "note", "remark")


class RowType(StrEnum):
    EMPTY = "EMPTY"
    NOTE = "NOTE"
    TOTAL = "TOTAL"
    WBS = "WBS"
    ITEM = "ITEM"
    UNKNOWN = "UNKNOWN"


class RowClassifier:
    """Classify rows from structure, not language-specific total keywords."""

    def classify(self, row: ExtractedRow) -> RowType:
        if not row.description:
            return RowType.EMPTY
        if self._is_total(row):
            return RowType.TOTAL
        if row.description.casefold().startswith(NOTE_WORDS):
            return RowType.NOTE
        if row.unit or any(v is not None for v in (row.qty, row.material, row.labor, row.amount)):
            return RowType.ITEM
        if numbering_level(row.item_no) or numbering_level(row.description) or row.is_bold:
            return RowType.WBS
        # R1 keeps legacy behaviour for compatibility. MS-2 review will expose
        # this uncertainty instead of forcing it directly into a final WBS.
        return RowType.WBS

    @staticmethod
    def _is_total(row: ExtractedRow) -> bool:
        return (
            bool(row.description)
            and row.amount is not None
            and is_blank_value(row.raw_unit)
            and is_blank_value(row.raw_qty)
        )
