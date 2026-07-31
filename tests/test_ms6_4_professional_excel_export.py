from pathlib import Path

from openpyxl import Workbook, load_workbook

from boq_studio.core.exporter import normalize_workbook
from boq_studio.core.tree import DraftTree, DraftTreeNode, TreeNodeType
from boq_studio.models import NormalizeRequest, SheetConfig


def _node(node_id, node_type, description, row=None, children=None):
    return DraftTreeNode(node_id, node_type, description, source_row=row, children=children or [])


def test_export_adds_excel_outline_and_professional_page_setup(tmp_path: Path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "BUILT-IN"
    ws.append(["No", "Description", "Unit", "Qty", "Material", "Labor", "Amount"])
    ws.append([1, "Item A", "lot", 1, 70, 30, 100])
    ws.append([2, "Item B", "lot", 1, 120, 80, 200])
    wb.save(source)

    item_a = _node("item:a", TreeNodeType.ITEM, "Item A", 2)
    item_b = _node("item:b", TreeNodeType.ITEM, "Item B", 3)
    wbs4 = _node("wbs:4", TreeNodeType.WBS, "Wall Work", children=[item_a, item_b])
    wbs3 = _node("wbs:3", TreeNodeType.WBS, "Lobby", children=[wbs4])
    root = _node("root", TreeNodeType.SHEET_ROOT, "BUILT-IN", children=[wbs3])
    tree = DraftTree("BUILT-IN", root)
    cfg = SheetConfig(
        sheet_name="BUILT-IN", header_row=1, start_row=2,
        item_no_col=1, description_cols=[2], unit_col=3, qty_col=4,
        material_col=5, labor_col=6, amount_col=7,
    )
    normalize_workbook(NormalizeRequest(source, output, [cfg], {"BUILT-IN": tree}))

    out = load_workbook(output, data_only=False)
    sheet = out["BUILT-IN"]
    assert [sheet.row_dimensions[r].outlineLevel for r in range(2, 7)] == [0, 1, 2, 3, 3]
    assert sheet.sheet_properties.outlinePr.summaryBelow is False
    assert sheet.freeze_panes == "A2"
    assert sheet.auto_filter.ref == "A1:K6"
    assert sheet.page_setup.orientation == "landscape"
    assert sheet.page_setup.fitToWidth == 1
    assert sheet.print_title_rows == "$1:$1"
    assert sheet.print_area == "'BUILT-IN'!$A$1:$K$6"
    assert sheet["A3"].font.bold is True
    assert sheet["A5"].font.bold is False
    assert sheet["I5"].value == 70
    assert sheet["J5"].value == 30
    assert sheet["K5"].value == 100

    project = out["Project"]
    assert [project.row_dimensions[r].outlineLevel for r in range(2, 7)] == [0, 1, 2, 3, 3]
    assert project.sheet_properties.outlinePr.summaryBelow is False
    out.close()


def test_delete_key_dispatches_to_selected_tree_rows_without_tk_runtime():
    from boq_studio.app import DraftTreeDialog

    class FakeTree:
        def __init__(self, selected):
            self._selected = selected

        def selection(self):
            return self._selected

    dialog = object.__new__(DraftTreeDialog)
    called = []
    dialog.tree = FakeTree(("item:1",))
    dialog._delete = lambda: called.append(True)

    assert dialog._on_delete_key() == "break"
    assert called == [True]

    called.clear()
    dialog.tree = FakeTree(())
    assert dialog._on_delete_key() == "break"
    assert called == []
