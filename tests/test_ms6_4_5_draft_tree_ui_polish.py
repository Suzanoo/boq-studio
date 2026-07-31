from boq_studio.core.tree import DraftTree, DraftTreeNode, TreeNodeType
from boq_studio.tree_ui import (
    ACTION_COMMANDS,
    EDIT_COMMANDS,
    CommandBarMode,
    command_bar_mode,
    editing_grid,
    summarize_tree,
)


def make_tree() -> DraftTree:
    root = DraftTreeNode("root", TreeNodeType.SHEET_ROOT, "BOQ")
    wbs = DraftTreeNode("wbs", TreeNodeType.WBS, "Structural")
    virtual = DraftTreeNode("virtual", TreeNodeType.VIRTUAL_WBS, "Created WBS")
    item_a = DraftTreeNode("item-a", TreeNodeType.ITEM, "Concrete")
    item_b = DraftTreeNode("item-b", TreeNodeType.ITEM, "Formwork")
    detail = DraftTreeNode("detail", TreeNodeType.DETAIL, "Remark")
    root.children.extend([wbs, virtual])
    wbs.children.extend([item_a, detail])
    virtual.children.append(item_b)
    return DraftTree("BOQ", root)


def test_ui_polish_preserves_every_existing_edit_command():
    assert tuple(spec.key for spec in EDIT_COMMANDS) == (
        "indent",
        "outdent",
        "move_up",
        "move_down",
        "toggle_type",
        "create_wbs",
        "edit",
        "delete",
    )
    assert len({spec.key for spec in EDIT_COMMANDS}) == len(EDIT_COMMANDS)


def test_primary_actions_remain_save_cancel_and_approve():
    assert tuple(spec.key for spec in ACTION_COMMANDS) == ("save", "cancel", "approve")
    assert ACTION_COMMANDS[-1].text == "Use this tree"


def test_responsive_breakpoint_is_deterministic_at_boundary():
    assert command_bar_mode(1079) == CommandBarMode.COMPACT
    assert command_bar_mode(1080) == CommandBarMode.WIDE
    assert command_bar_mode(1600) == CommandBarMode.WIDE


def test_wide_layout_keeps_all_edit_commands_in_one_row():
    positions = editing_grid(CommandBarMode.WIDE, len(EDIT_COMMANDS))
    assert positions == tuple((0, index) for index in range(len(EDIT_COMMANDS)))


def test_compact_layout_wraps_commands_without_hiding_or_reordering():
    positions = editing_grid(CommandBarMode.COMPACT, len(EDIT_COMMANDS))
    assert positions == (
        (0, 0), (0, 1), (0, 2), (0, 3),
        (1, 0), (1, 1), (1, 2), (1, 3),
    )
    assert len(set(positions)) == len(EDIT_COMMANDS)


def test_tree_summary_excludes_sheet_root_and_counts_all_node_types():
    summary = summarize_tree(make_tree())
    assert summary.nodes == 5
    assert summary.wbs == 2
    assert summary.items == 2
    assert summary.other == 1
    assert summary.primary_line == "5 nodes"
    assert summary.secondary_line == "2 WBS  •  2 Items  •  1 Other"


def test_tree_summary_omits_other_when_zero():
    root = DraftTreeNode("root", TreeNodeType.SHEET_ROOT, "BOQ")
    root.children.append(DraftTreeNode("item", TreeNodeType.ITEM, "Concrete"))
    summary = summarize_tree(DraftTree("BOQ", root))
    assert summary.secondary_line == "0 WBS  •  1 Items"


def test_ui_polish_does_not_claim_enter_for_approve_action():
    # Enter already edits a selected description in DraftTreeDialog. Keeping the
    # approve shortcut empty prevents UI polish from silently changing that contract.
    approve = next(spec for spec in ACTION_COMMANDS if spec.key == "approve")
    edit = next(spec for spec in EDIT_COMMANDS if spec.key == "edit")
    assert approve.shortcut == ""
    assert "Enter" in edit.shortcut
