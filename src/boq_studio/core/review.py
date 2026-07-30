from __future__ import annotations

from dataclasses import dataclass

from .classification import ClassificationResult, RowClassifier, RowType, resolve_row_type
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
    auto_type: RowType
    overridden: bool


class ClassificationReviewService:
    """Prepare classifier decisions for the user-review workspace."""

    def __init__(self, classifier: RowClassifier | None = None) -> None:
        self.classifier = classifier or RowClassifier()
        self.extractor = SheetExtractor()

    def build_rows(self, ws, cfg, include_empty: bool = False) -> list[ReviewRow]:
        review_rows: list[ReviewRow] = []
        overrides = getattr(cfg, "classification_overrides", {})
        for extracted in self.extractor.extract(ws, cfg):
            decision = self.classifier.classify_result(extracted)
            effective = resolve_row_type(decision, extracted.source_row, overrides)
            if not include_empty and effective == RowType.EMPTY:
                continue
            review_rows.append(self._to_review_row(extracted, decision, effective))
        return review_rows

    @staticmethod
    def _to_review_row(
        row: ExtractedRow,
        decision: ClassificationResult,
        effective: RowType,
    ) -> ReviewRow:
        return ReviewRow(
            source_row=row.source_row,
            item_no=row.item_no,
            description=row.description,
            row_type=effective,
            confidence=decision.confidence,
            reason=decision.reason,
            unit=row.unit,
            qty=row.qty,
            amount=row.amount,
            auto_type=decision.row_type,
            overridden=effective != decision.row_type,
        )
