from __future__ import annotations

from dataclasses import dataclass

from .classification import ClassificationResult, RowClassifier, RowType
from .extraction import ExtractedRow, SheetExtractor


@dataclass(slots=True, frozen=True)
class ReviewRow:
    source_row: int
    item_no: str
    description: str
    row_type: RowType
    confidence: float
    reason: str
    unit: str
    qty: float | None
    amount: float | None


class ClassificationReviewService:
    """Prepare classifier decisions for a future user-review workspace."""

    def __init__(self, classifier: RowClassifier | None = None) -> None:
        self.classifier = classifier or RowClassifier()
        self.extractor = SheetExtractor()

    def build_rows(self, ws, cfg, include_empty: bool = False) -> list[ReviewRow]:
        review_rows: list[ReviewRow] = []
        for extracted in self.extractor.extract(ws, cfg):
            decision = self.classifier.classify_result(extracted)
            if not include_empty and decision.row_type == RowType.EMPTY:
                continue
            review_rows.append(self._to_review_row(extracted, decision))
        return review_rows

    @staticmethod
    def _to_review_row(row: ExtractedRow, decision: ClassificationResult) -> ReviewRow:
        return ReviewRow(
            source_row=row.source_row,
            item_no=row.item_no,
            description=row.description,
            row_type=decision.row_type,
            confidence=decision.confidence,
            reason=decision.reason,
            unit=row.unit,
            qty=row.qty,
            amount=row.amount,
        )
