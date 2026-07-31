from pathlib import Path

from boq_studio.core.tree import DraftTree, DraftTreeNode, TreeNodeType
from boq_studio.models import SheetConfig
from boq_studio.session import (
    build_session_payload,
    config_from_dict,
    load_session,
    save_session,
    session_matches_workbook,
    session_path_for,
    tree_from_dict,
)
from boq_studio.workflow import SheetState, sheet_state, workflow_status


def sample_tree() -> DraftTree:
    root = DraftTreeNode("sheet:ARCH", TreeNodeType.SHEET_ROOT, "ARCH")
    wbs = DraftTreeNode("ARCH:wbs", TreeNodeType.WBS, "Concrete Work")
    item = DraftTreeNode("ARCH:12", TreeNodeType.ITEM, "Concrete 280 ksc", source_row=12, amount=1250.0)
    wbs.children.append(item)
    root.children.append(wbs)
    return DraftTree("ARCH", root)


def test_tree_draft_blocks_export_until_approved():
    assert sheet_state(has_mapping=True, has_tree=True) is SheetState.TREE_DRAFT
    assert sheet_state(has_mapping=True, has_tree=True, tree_ready=True) is SheetState.TREE_READY
    draft = workflow_status([SheetState.TREE_DRAFT])
    assert draft.tree_draft_count == 1
    assert not draft.can_export
    assert workflow_status([SheetState.TREE_READY]).can_export


def test_session_round_trip_preserves_mapping_tree_and_ready_state(tmp_path: Path):
    workbook = tmp_path / "sample.xlsx"
    workbook.write_bytes(b"workbook-v1")
    config = SheetConfig(
        sheet_name="ARCH",
        header_row=4,
        start_row=6,
        description_cols=[2, 3],
        amount_col=9,
        classification_overrides={12: "ITEM"},
        wbs_level_overrides={8: 3},
    )
    payload = build_session_payload(
        workbook,
        {"ARCH": True},
        {"ARCH": config},
        {"ARCH": sample_tree()},
        {"ARCH"},
        set(),
    )
    path = session_path_for(workbook, tmp_path / "sessions")
    save_session(path, payload)
    restored = load_session(path)

    restored_config = config_from_dict(restored["configs"]["ARCH"])
    restored_tree = tree_from_dict(restored["trees"]["ARCH"])
    assert restored_config.classification_overrides == {12: "ITEM"}
    assert restored_config.wbs_level_overrides == {8: 3}
    assert restored_tree.root.children[0].children[0].amount == 1250.0
    assert restored["tree_ready"] == ["ARCH"]
    assert session_matches_workbook(restored, workbook)


def test_changed_workbook_rejects_saved_session(tmp_path: Path):
    workbook = tmp_path / "sample.xlsx"
    workbook.write_bytes(b"workbook-v1")
    payload = build_session_payload(workbook, {}, {}, {}, set(), set())
    workbook.write_bytes(b"workbook-v2")
    assert not session_matches_workbook(payload, workbook)
