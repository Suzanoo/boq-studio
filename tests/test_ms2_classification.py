from openpyxl import Workbook

from boq_studio.core.classification import RowClassifier, RowType
from boq_studio.core.extraction import ExtractedRow
from boq_studio.core.review import ClassificationReviewService
from boq_studio.models import SheetConfig


def make_row(**overrides):
    values = dict(
        source_row=2,
        item_no="",
        description="Example",
        unit="",
        qty=None,
        material=None,
        labor=None,
        amount=None,
        raw_unit=None,
        raw_qty=None,
        is_bold=False,
        indent_level=1,
    )
    values.update(overrides)
    return ExtractedRow(**values)


def test_structural_total_is_language_independent():
    decision = RowClassifier().classify_result(make_row(description="Anything", amount=100.0))
    assert decision.row_type == RowType.TOTAL
    assert decision.confidence >= 0.99


def test_item_with_total_word_is_not_total_when_unit_or_qty_exists():
    decision = RowClassifier().classify_result(
        make_row(description="AIR SPLIT TYPE รวมอุปกรณ์", unit="set", qty=1.0, amount=100.0,
                 raw_unit="set", raw_qty=1)
    )
    assert decision.row_type == RowType.ITEM


def test_numbered_heading_is_wbs():
    decision = RowClassifier().classify_result(make_row(item_no="1.2", description="Substructure"))
    assert decision.row_type == RowType.WBS
    assert decision.confidence >= 0.9


def test_ambiguous_text_only_row_is_unknown():
    decision = RowClassifier().classify_result(make_row(description="General works"))
    assert decision.row_type == RowType.UNKNOWN
    assert decision.confidence < 0.5


def test_bold_text_only_row_is_probable_wbs():
    decision = RowClassifier().classify_result(make_row(description="Structural works", is_bold=True))
    assert decision.row_type == RowType.WBS
    assert 0.5 < decision.confidence < 0.9


def test_review_service_exposes_reason_and_confidence():
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"
    ws.append(["Item", "Description", "Unit", "Qty", "Amount"])
    ws.append([None, "General works", None, None, None])
    ws.append([None, "Concrete", "m3", 10, 1000])
    cfg = SheetConfig(
        sheet_name="BOQ",
        start_row=2,
        item_no_col=1,
        description_cols=[2],
        unit_col=3,
        qty_col=4,
        amount_col=5,
    )
    rows = ClassificationReviewService().build_rows(ws, cfg)
    assert rows[0].row_type == RowType.UNKNOWN
    assert rows[0].reason
    assert rows[1].row_type == RowType.ITEM
