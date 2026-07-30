from boq_studio.core.tree import DraftTree, DraftTreeNode, TreeNodeType
from boq_studio.core.tree_edit import DraftTreeEditor


def node(node_id: str, description: str, kind=TreeNodeType.ITEM, children=None):
    return DraftTreeNode(
        node_id=node_id,
        node_type=kind,
        description=description,
        children=list(children or []),
    )


def make_tree(children):
    root = node("root", "ARCH", TreeNodeType.SHEET_ROOT, children)
    return DraftTree(sheet_name="ARCH", root=root)


def test_edit_description_for_item_and_wbs():
    wbs = node("w", "Old WBS", TreeNodeType.WBS)
    item = node("i", "Old item")
    editor = DraftTreeEditor(make_tree([wbs, item]))

    editor.rename_node("w", "New WBS")
    editor.rename_node("i", "New item")

    assert wbs.description == "New WBS"
    assert item.description == "New item"


def test_delete_multiple_rows_promotes_children():
    child = node("child", "Child")
    heading = node("heading", "Heading", TreeNodeType.WBS, [child])
    other = node("other", "Other")
    final = node("final", "Final")
    tree = make_tree([heading, other, final])
    editor = DraftTreeEditor(tree)

    editor.delete_nodes(["heading", "other"])

    assert [n.node_id for n in tree.root.children] == ["child", "final"]


def test_move_selected_block_up_preserves_order():
    tree = make_tree([node("a", "A"), node("b", "B"), node("c", "C"), node("d", "D")])
    editor = DraftTreeEditor(tree)

    editor.move_nodes_up(["b", "c"])

    assert [n.node_id for n in tree.root.children] == ["b", "c", "a", "d"]


def test_move_selected_block_down_preserves_order():
    tree = make_tree([node("a", "A"), node("b", "B"), node("c", "C"), node("d", "D")])
    editor = DraftTreeEditor(tree)

    editor.move_nodes_down(["b", "c"])

    assert [n.node_id for n in tree.root.children] == ["a", "d", "b", "c"]


def test_delete_selected_parent_and_child_removes_both():
    child = node("child", "Child")
    heading = node("heading", "Heading", TreeNodeType.WBS, [child])
    final = node("final", "Final")
    tree = make_tree([heading, final])
    editor = DraftTreeEditor(tree)

    editor.delete_nodes(["heading", "child"])

    assert [n.node_id for n in tree.root.children] == ["final"]


def test_toggle_single_item_to_wbs_and_back_to_item():
    item = node("i", "Item")
    editor = DraftTreeEditor(make_tree([item]))

    assert editor.toggle_wbs_item(["i"]) == TreeNodeType.WBS
    assert item.node_type == TreeNodeType.WBS

    assert editor.toggle_wbs_item(["i"]) == TreeNodeType.ITEM
    assert item.node_type == TreeNodeType.ITEM


def test_toggle_all_wbs_selection_to_items():
    first = node("a", "A", TreeNodeType.WBS)
    second = node("b", "B", TreeNodeType.WBS)
    editor = DraftTreeEditor(make_tree([first, second]))

    editor.toggle_wbs_item(["a", "b"])

    assert first.node_type == TreeNodeType.ITEM
    assert second.node_type == TreeNodeType.ITEM


def test_toggle_mixed_selection_to_wbs():
    first = node("a", "A", TreeNodeType.WBS)
    second = node("b", "B", TreeNodeType.ITEM)
    editor = DraftTreeEditor(make_tree([first, second]))

    editor.toggle_wbs_item(["a", "b"])

    assert first.node_type == TreeNodeType.WBS
    assert second.node_type == TreeNodeType.WBS
