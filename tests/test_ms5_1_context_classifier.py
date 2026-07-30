from boq_studio.core.classification import RowClassifier, RowType
from boq_studio.core.extraction import ExtractedRow


def row(source_row: int, **values) -> ExtractedRow:
    data = dict(
        source_row=source_row,
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
    data.update(values)
    return ExtractedRow(**data)


def test_numbered_summary_header_with_children_is_wbs_not_total():
    rows = [
        row(1, item_no="1", description="High voltage system", amount=1000),
        row(2, item_no="1.1", description="HV cable", unit="m", qty=10, raw_unit="m", raw_qty=10),
        row(3, item_no="1.2", description="Concrete foundation", unit="set", qty=1, raw_unit="set", raw_qty=1),
    ]
    decisions = RowClassifier().classify_rows(rows)
    assert decisions[0].row_type == RowType.WBS
    assert decisions[0].confidence >= 0.95


def test_closing_amount_row_after_items_is_total():
    rows = [
        row(1, item_no="1", description="Floor work"),
        row(2, item_no="1.1", description="Floor A", unit="m2", qty=10, raw_unit="m2", raw_qty=10),
        row(3, item_no="1.2", description="Floor B", unit="m2", qty=20, raw_unit="m2", raw_qty=20),
        row(4, description="Floor work summary", amount=3000),
        row(5, item_no="2", description="Wall work", is_bold=True),
    ]
    decisions = RowClassifier().classify_rows(rows)
    assert decisions[3].row_type == RowType.TOTAL


def test_bold_amount_heading_before_items_is_wbs():
    rows = [
        row(1, item_no="1", description="Lobby work", amount=5000, is_bold=True),
        row(2, description="Glass wall", unit="m2", qty=11, raw_unit="m2", raw_qty=11),
    ]
    decisions = RowClassifier().classify_rows(rows)
    assert decisions[0].row_type == RowType.WBS


def test_dash_text_without_quantity_is_detail_not_wbs():
    rows = [row(1, description="- Steel Plate THK 5 mm.")]
    decision = RowClassifier().classify_rows(rows)[0]
    assert decision.row_type == RowType.DETAIL


def test_item_no_does_not_override_item_structure():
    rows = [row(1, item_no="1.1", description="Cable", unit="m", qty=9, raw_unit="m", raw_qty=9)]
    decision = RowClassifier().classify_rows(rows)[0]
    assert decision.row_type == RowType.ITEM
