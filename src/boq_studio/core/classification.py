from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
import re
from collections.abc import Sequence

from .extraction import ExtractedRow
from .values import is_blank_value, numbering_level

NOTE_WORDS = ("หมายเหตุ", "note", "remark")


class RowType(StrEnum):
    EMPTY = "EMPTY"
    NOTE = "NOTE"
    TOTAL = "TOTAL"
    WBS = "WBS"
    ITEM = "ITEM"
    DETAIL = "DETAIL"
    UNKNOWN = "UNKNOWN"
    IGNORE = "IGNORE"


@dataclass(slots=True, frozen=True)
class ClassificationResult:
    row_type: RowType
    confidence: float
    reason: str


def resolve_row_type(
    decision: ClassificationResult,
    source_row: int,
    overrides: dict[int, str] | None = None,
) -> RowType:
    """Return a reviewed row type when a valid override exists."""
    if not overrides:
        return decision.row_type
    raw = overrides.get(source_row)
    if raw is None:
        return decision.row_type
    try:
        return RowType(raw)
    except ValueError:
        return decision.row_type


class RowClassifier:
    """Classify BOQ rows using both row structure and neighbouring rows.

    ``classify_result`` remains as a compatibility API for single-row callers.
    New workflows should call ``classify_rows`` because BOQ headings and totals
    cannot be identified reliably without looking at rows before and after them.
    """

    def classify(self, row: ExtractedRow) -> RowType:
        return self.classify_result(row).row_type

    def classify_result(self, row: ExtractedRow) -> ClassificationResult:
        """Classify one isolated row using only strong local signals."""
        if not row.description:
            return ClassificationResult(RowType.EMPTY, 1.00, "Description is blank")

        if row.description.casefold().startswith(NOTE_WORDS):
            return ClassificationResult(RowType.NOTE, 0.95, "Description starts with a note marker")

        if row.unit or row.qty is not None:
            return ClassificationResult(RowType.ITEM, 0.99, "Unit or Qty is present")

        if any(value is not None for value in (row.material, row.labor)):
            return ClassificationResult(RowType.ITEM, 0.90, "Material or Labor value is present")

        if self._looks_like_detail(row):
            return ClassificationResult(RowType.DETAIL, 0.70, "Text-only row looks like item detail")

        if self._has_summary_shape(row):
            return ClassificationResult(
                RowType.TOTAL,
                0.99,
                "Description and Amount are present while Unit and Qty are blank",
            )

        explicit_level = numbering_level(row.item_no) or numbering_level(row.description)
        if explicit_level is not None:
            return ClassificationResult(RowType.WBS, 0.95, f"Detected numbering level {explicit_level}")

        if row.is_bold:
            return ClassificationResult(RowType.WBS, 0.65, "Text-only row is bold")

        if row.indent_level > 1:
            return ClassificationResult(RowType.UNKNOWN, 0.40, "Indented text-only row needs context")

        return ClassificationResult(
            RowType.UNKNOWN,
            0.35,
            "Text-only row has no reliable item, total, note, numbering, or formatting signal",
        )

    def classify_rows(self, rows: Sequence[ExtractedRow]) -> list[ClassificationResult]:
        """Classify a sheet sequence using Item No. hierarchy and row position."""
        results: list[ClassificationResult] = []
        for index, row in enumerate(rows):
            local = self.classify_result(row)
            if local.row_type in {RowType.EMPTY, RowType.NOTE, RowType.ITEM, RowType.DETAIL}:
                results.append(local)
                continue

            if self._has_numbered_children(rows, index):
                results.append(
                    ClassificationResult(
                        RowType.WBS,
                        0.98,
                        "Item No. is the parent of following numbered rows",
                    )
                )
                continue

            if self._is_probable_heading_before_items(rows, index):
                results.append(
                    ClassificationResult(
                        RowType.WBS,
                        0.88,
                        "Heading appears before one or more item rows",
                    )
                )
                continue

            if self._has_summary_shape(row):
                if self._is_closing_total(rows, index):
                    results.append(
                        ClassificationResult(
                            RowType.TOTAL,
                            0.92,
                            "Amount-only row closes a preceding item group",
                        )
                    )
                else:
                    results.append(
                        ClassificationResult(
                            RowType.UNKNOWN,
                            0.45,
                            "Amount-only row has no clear child or closing-group context",
                        )
                    )
                continue

            results.append(local)
        return results

    @staticmethod
    def _has_summary_shape(row: ExtractedRow) -> bool:
        return (
            bool(row.description)
            and row.amount is not None
            and is_blank_value(row.raw_unit)
            and is_blank_value(row.raw_qty)
        )

    @staticmethod
    def _looks_like_detail(row: ExtractedRow) -> bool:
        text = row.description.lstrip()
        return (
            not row.item_no
            and row.amount is None
            and not row.unit
            and row.qty is None
            and bool(re.match(r"^[-–—•*()]", text))
        )

    def _has_numbered_children(self, rows: Sequence[ExtractedRow], index: int) -> bool:
        parent = self._number_path(rows[index].item_no)
        if parent is None:
            return False
        for following in rows[index + 1 : index + 10]:
            if not following.description:
                continue
            child = self._number_path(following.item_no)
            if child is None:
                if following.unit or following.qty is not None:
                    continue
                break
            if len(child) <= len(parent):
                break
            return child[: len(parent)] == parent
        return False

    def _is_probable_heading_before_items(self, rows: Sequence[ExtractedRow], index: int) -> bool:
        row = rows[index]
        if row.unit or row.qty is not None or any(v is not None for v in (row.material, row.labor)):
            return False
        if self._looks_like_detail(row):
            return False
        seen_items = 0
        for following in rows[index + 1 : index + 7]:
            if not following.description:
                continue
            if following.unit or following.qty is not None:
                seen_items += 1
                if seen_items >= 1:
                    return bool(row.item_no or row.is_bold or row.amount is not None)
                continue
            if self._has_summary_shape(following):
                continue
            if numbering_level(following.item_no) is not None:
                continue
            break
        return False

    def _is_closing_total(self, rows: Sequence[ExtractedRow], index: int) -> bool:
        row = rows[index]
        if not self._has_summary_shape(row):
            return False
        # A numbered row followed by descendants is a group header, never a closing total.
        if self._has_numbered_children(rows, index):
            return False

        previous_items = 0
        for previous in reversed(rows[max(0, index - 8) : index]):
            if not previous.description:
                continue
            if previous.unit or previous.qty is not None:
                previous_items += 1
                continue
            if self._looks_like_detail(previous):
                continue
            if previous_items:
                break
        if previous_items == 0:
            return False

        # A following sibling/header indicates this row closes the previous group.
        for following in rows[index + 1 : index + 5]:
            if not following.description:
                continue
            if numbering_level(following.item_no) is not None:
                return True
            if following.is_bold and not (following.unit or following.qty is not None):
                return True
            break
        # End-of-sheet totals are also closing totals.
        return not any(f.description for f in rows[index + 1 :])

    @staticmethod
    def _number_path(text: str) -> tuple[int, ...] | None:
        value = text.strip().rstrip(")")
        if not re.fullmatch(r"\d+(?:\.\d+){0,5}", value):
            return None
        return tuple(int(part) for part in value.split("."))
