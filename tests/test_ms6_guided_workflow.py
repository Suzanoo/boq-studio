from boq_studio.workflow import SheetState, sheet_state, workflow_status


def test_sheet_state_progression():
    assert sheet_state(has_mapping=False, has_tree=False) is SheetState.NEW
    assert sheet_state(has_mapping=True, has_tree=False) is SheetState.MAPPED
    assert sheet_state(has_mapping=True, has_tree=True) is SheetState.TREE_DRAFT
    assert sheet_state(has_mapping=True, has_tree=True, tree_ready=True) is SheetState.TREE_READY
    assert sheet_state(has_mapping=True, has_tree=True, exported=True) is SheetState.EXPORTED


def test_export_locked_until_every_selected_sheet_has_tree():
    status = workflow_status([SheetState.TREE_READY, SheetState.MAPPED])
    assert status.can_export is False
    assert status.mapped_count == 1


def test_export_enabled_when_all_selected_sheets_are_tree_ready():
    status = workflow_status([SheetState.TREE_READY, SheetState.TREE_READY])
    assert status.can_export is True
    assert status.summary == "2 selected · Ready to normalize"


def test_exported_sheets_still_satisfy_workflow():
    status = workflow_status([SheetState.EXPORTED, SheetState.TREE_READY])
    assert status.can_export is True


def test_no_selection_cannot_export():
    status = workflow_status([])
    assert status.can_export is False
    assert status.summary == "No breakdown sheets selected"
