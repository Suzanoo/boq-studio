from __future__ import annotations

from dataclasses import dataclass

from .classification import ClassificationResult, RowClassifier, RowType, resolve_row_type
from .extraction import ExtractedRow, SheetExtractor
from .wbs import infer_wbs_level


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
    auto_wbs_level: int | None
    wbs_level: int | None
    level_overridden: bool


class ClassificationReviewService:
    """Prepare row-type and WBS-level decisions for user review."""

    def __init__(self, classifier: RowClassifier | None = None) -> None:
        self.classifier = classifier or RowClassifier()
        self.extractor = SheetExtractor()

    def build_rows(self, ws, cfg, include_empty: bool = False) -> list[ReviewRow]:
        review_rows: list[ReviewRow] = []
        overrides = getattr(cfg, "classification_overrides", {})
        level_overrides = getattr(cfg, "wbs_level_overrides", {})
        for extracted in self.extractor.extract(ws, cfg):
            decision = self.classifier.classify_result(extracted)
            effective = resolve_row_type(decision, extracted.source_row, overrides)
            if not include_empty and effective == RowType.EMPTY:
                continue
            review_rows.append(
                self._to_review_row(extracted, decision, effective, level_overrides)
            )
        return review_rows

    @staticmethod
    def _to_review_row(
        row: ExtractedRow,
        decision: ClassificationResult,
        effective: RowType,
        level_overrides: dict[int, int],
    ) -> ReviewRow:
        auto_level = infer_wbs_level(row) if decision.row_type == RowType.WBS else None
        raw_override = level_overrides.get(row.source_row)
        final_level = raw_override if effective == RowType.WBS else None
        if final_level is None and effective == RowType.WBS:
            final_level = auto_level or 3
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
            auto_wbs_level=auto_level,
            wbs_level=final_level,
            level_overridden=raw_override is not None,
        )
