import pytest

from boq_studio.core.tree import DraftTree, DraftTreeNode, TreeNodeType
from boq_studio.core.tree_edit import DraftTreeEditor, TreeEditError


def node(node_id, kind, description, children=None):
    return DraftTreeNode(node_id, kind, description, children=children or [])


def tree_with(*children):
    return DraftTree("ARCH", node("sheet:ARCH", TreeNodeType.SHEET_ROOT, "ARCH", list(children)))


def test_indent_moves_row_under_previous_wbs():
    heading = node("w1", TreeNodeType.WBS, "Floor")
    item = node("i1", TreeNodeType.ITEM, "Concrete")
    tree = tree_with(heading, item)
    DraftTreeEditor(tree).indent("i1")
    assert tree.root.children == [heading]
    assert heading.children == [item]


def test_indent_rejects_previous_item():
    tree = tree_with(node("i1", TreeNodeType.ITEM, "A"), node("i2", TreeNodeType.ITEM, "B"))
    with pytest.raises(TreeEditError):
        DraftTreeEditor(tree).indent("i2")


def test_outdent_places_node_after_parent():
    item = node("i1", TreeNodeType.ITEM, "Concrete")
    heading = node("w1", TreeNodeType.WBS, "Floor", [item])
    other = node("w2", TreeNodeType.WBS, "Wall")
    tree = tree_with(heading, other)
    DraftTreeEditor(tree).outdent("i1")
    assert tree.root.children == [heading, item, other]


def test_create_virtual_wbs_under_selected_wbs():
    heading = node("w1", TreeNodeType.WBS, "Architecture")
    tree = tree_with(heading)
    created = DraftTreeEditor(tree).create_wbs("Floor", "w1")
    assert created.node_type == TreeNodeType.VIRTUAL_WBS
    assert heading.children == [created]


def test_delete_virtual_wbs_promotes_children():
    item = node("i1", TreeNodeType.ITEM, "Concrete")
    virtual = node("v1", TreeNodeType.VIRTUAL_WBS, "Floor", [item])
    tree = tree_with(virtual)
    DraftTreeEditor(tree).delete_virtual_wbs("v1")
    assert tree.root.children == [item]


def test_source_wbs_cannot_be_deleted():
    tree = tree_with(node("w1", TreeNodeType.WBS, "Floor"))
    with pytest.raises(TreeEditError):
        DraftTreeEditor(tree).delete_virtual_wbs("w1")


def test_rename_wbs_but_not_item():
    heading = node("w1", TreeNodeType.WBS, "Old")
    item = node("i1", TreeNodeType.ITEM, "Concrete")
    tree = tree_with(heading, item)
    editor = DraftTreeEditor(tree)
    editor.rename_wbs("w1", "New")
    assert heading.description == "New"
    with pytest.raises(TreeEditError):
        editor.rename_wbs("i1", "No")
