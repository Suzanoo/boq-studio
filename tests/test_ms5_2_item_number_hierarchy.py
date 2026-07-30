from boq_studio.core.classification import RowClassifier, RowType
from boq_studio.core.extraction import ExtractedRow
from boq_studio.core.item_numbering import ItemNumberParser, NumberingStyle


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


def test_parser_supports_numeric_hierarchy_and_wrappers():
    parser = ItemNumberParser()
    assert parser.parse("1").path == ("1",)
    assert parser.parse("1.2.03").path == ("1", "2", "3")
    assert parser.parse("1)").path == ("1",)
    assert parser.parse("(1.2)").path == ("1", "2")


def test_parser_supports_alpha_hierarchy():
    number = ItemNumberParser().parse("a.2.1")
    assert number is not None
    assert number.style == NumberingStyle.ALPHA
    assert number.path == ("A", "2", "1")


def test_parser_rejects_product_and_model_codes():
    parser = ItemNumberParser()
    for value in ("F1.1", "FDX-301", "C1.2", "M-01", "1/2"):
        assert parser.parse(value) is None


def test_parent_and_sibling_relationships():
    parser = ItemNumberParser()
    parent = parser.parse("2")
    child = parser.parse("2.1")
    sibling = parser.parse("2.2")
    assert parent is not None and child is not None and sibling is not None
    assert parent.is_parent_of(child)
    assert child.is_sibling_of(sibling)


def test_alpha_parent_with_numbered_children_is_detected_as_wbs():
    rows = [
        row(1, item_no="A", description="Architectural work", amount=1000),
        row(2, item_no="A.1", description="Wall item", unit="m2", qty=10, raw_unit="m2", raw_qty=10),
        row(3, item_no="A.2", description="Floor item", unit="m2", qty=20, raw_unit="m2", raw_qty=20),
    ]
    assert RowClassifier().classify_rows(rows)[0].row_type == RowType.WBS


def test_product_code_does_not_create_false_numbered_parent():
    rows = [
        row(1, item_no="F1.1", description="Floor finish heading", amount=1000),
        row(2, item_no="F1.1.1", description="Product code child", unit="set", qty=1, raw_unit="set", raw_qty=1),
    ]
    decisions = RowClassifier().classify_rows(rows)
    assert "parent of following numbered rows" not in decisions[0].reason
