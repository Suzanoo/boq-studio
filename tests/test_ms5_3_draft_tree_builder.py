from boq_studio.core.classification import RowClassifier
from boq_studio.core.extraction import ExtractedRow
from boq_studio.core.tree import DraftTreeBuilder, TreeNodeType


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


def build(rows):
    decisions = RowClassifier().classify_rows(rows)
    return DraftTreeBuilder().build("ARCH", rows, decisions)


def test_numbered_heading_owns_following_numbered_items():
    rows = [
        row(31, item_no="1", description="Floor work", amount=1000),
        row(32, item_no="1.1", description="Floor A", unit="m2", qty=10, raw_unit="m2", raw_qty=10),
        row(33, item_no="1.2", description="Floor B", unit="m2", qty=20, raw_unit="m2", raw_qty=20),
    ]
    tree = build(rows)
    heading = tree.root.children[0]
    assert heading.node_type == TreeNodeType.WBS
    assert [child.description for child in heading.children] == ["Floor A", "Floor B"]


def test_nested_numbered_wbs_uses_item_number_ancestry():
    rows = [
        row(1, item_no="1", description="Architecture", is_bold=True),
        row(2, item_no="1.1", description="Floor work", is_bold=True),
        row(3, item_no="1.1.1", description="Floor item", unit="m2", qty=5, raw_unit="m2", raw_qty=5),
        row(4, item_no="2", description="Wall work", is_bold=True),
        row(5, item_no="2.1", description="Wall item", unit="m2", qty=6, raw_unit="m2", raw_qty=6),
    ]
    tree = build(rows)
    architecture, wall = tree.root.children
    assert architecture.description == "Architecture"
    assert architecture.children[0].description == "Floor work"
    assert architecture.children[0].children[0].description == "Floor item"
    assert wall.description == "Wall work"
    assert wall.children[0].description == "Wall item"


def test_detail_row_attaches_to_previous_item():
    rows = [
        row(1, item_no="1", description="Built-in work", is_bold=True),
        row(2, item_no="1.1", description="Glass wall", unit="m2", qty=1, raw_unit="m2", raw_qty=1),
        row(3, description="- Stainless steel frame"),
    ]
    tree = build(rows)
    item = tree.root.children[0].children[0]
    assert item.node_type == TreeNodeType.ITEM
    assert item.children[0].node_type == TreeNodeType.DETAIL


def test_unknown_rows_are_preserved_for_future_tree_review():
    rows = [row(1, description="Ambiguous text")]
    tree = build(rows)
    assert tree.root.children[0].node_type == TreeNodeType.UNKNOWN
    assert tree.root.children[0].source_row == 1


def test_override_can_promote_unknown_row_to_wbs():
    rows = [
        row(1, description="User-created heading"),
        row(2, description="Item", unit="set", qty=1, raw_unit="set", raw_qty=1),
    ]
    decisions = RowClassifier().classify_rows(rows)
    tree = DraftTreeBuilder().build("ARCH", rows, decisions, overrides={1: "WBS"})
    assert tree.root.children[0].node_type == TreeNodeType.WBS
    assert tree.root.children[0].children[0].description == "Item"


def test_walk_returns_depths_for_tree_editor():
    rows = [
        row(1, item_no="1", description="Floor work", is_bold=True),
        row(2, item_no="1.1", description="Floor item", unit="m2", qty=1, raw_unit="m2", raw_qty=1),
    ]
    walked = [(depth, node.description) for depth, node in build(rows).walk()]
    assert walked == [(1, "Floor work"), (2, "Floor item")]
