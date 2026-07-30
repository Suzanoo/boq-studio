from __future__ import annotations

from dataclasses import dataclass
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


@dataclass(slots=True, frozen=True)
class ClassificationResult:
    row_type: RowType
    confidence: float
    reason: str


class RowClassifier:
    """Classify BOQ rows from structure and formatting signals.

    The classifier does not mutate WBS state and does not rely on language-specific
    total keywords. Ambiguous text-only rows are exposed as UNKNOWN so a review
    workflow can resolve them later.
    """

    def classify(self, row: ExtractedRow) -> RowType:
        """Compatibility API used by the current exporter."""
        return self.classify_result(row).row_type

    def classify_result(self, row: ExtractedRow) -> ClassificationResult:
        if not row.description:
            return ClassificationResult(RowType.EMPTY, 1.00, "Description is blank")

        if self._is_total(row):
            return ClassificationResult(
                RowType.TOTAL,
                0.99,
                "Description and Amount are present while Unit and Qty are blank",
            )

        if row.description.casefold().startswith(NOTE_WORDS):
            return ClassificationResult(RowType.NOTE, 0.95, "Description starts with a note marker")

        if row.unit or row.qty is not None:
            return ClassificationResult(RowType.ITEM, 0.99, "Unit or Qty is present")

        if any(value is not None for value in (row.material, row.labor)):
            return ClassificationResult(RowType.ITEM, 0.90, "Material or Labor value is present")

        explicit_level = numbering_level(row.item_no) or numbering_level(row.description)
        if explicit_level is not None:
            return ClassificationResult(RowType.WBS, 0.95, f"Detected numbering level {explicit_level}")

        if row.is_bold:
            return ClassificationResult(RowType.WBS, 0.75, "Text-only row is bold")

        if row.indent_level > 1:
            return ClassificationResult(RowType.WBS, 0.65, "Text-only row is indented")

        return ClassificationResult(
            RowType.UNKNOWN,
            0.35,
            "Text-only row has no reliable item, total, note, numbering, or formatting signal",
        )

    @staticmethod
    def _is_total(row: ExtractedRow) -> bool:
        return (
            bool(row.description)
            and row.amount is not None
            and is_blank_value(row.raw_unit)
            and is_blank_value(row.raw_qty)
        )
