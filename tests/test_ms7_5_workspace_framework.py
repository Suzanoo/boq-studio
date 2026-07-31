import pytest

from boq_studio.commands import CommandRegistry
from boq_studio.workspace import (
    WorkspaceKind,
    WorkspaceLifecycle,
    WorkspaceServices,
    WorkspaceSpec,
    workspace_surface_contract,
)


def spec(workspace_id="draft_tree", **changes):
    values = dict(
        workspace_id=workspace_id,
        kind=WorkspaceKind.DRAFT_TREE,
        title="Draft WBS tree",
        geometry="1050x760",
        min_size=(760, 520),
        surfaces=("menu", "toolbar", "workspace", "sidebar", "command_bar", "status_bar"),
    )
    values.update(changes)
    return WorkspaceSpec(**values)


def test_workspace_spec_normalizes_identity_and_deduplicates_surfaces():
    actual = spec(workspace_id="  draft_tree  ", surfaces=("workspace", "sidebar", "sidebar")).normalized()
    assert actual.workspace_id == "draft_tree"
    assert actual.surfaces == ("workspace", "sidebar")


def test_workspace_spec_rejects_invalid_identity_and_minimum_size():
    with pytest.raises(ValueError):
        spec(workspace_id=" ").normalized()
    with pytest.raises(ValueError):
        spec(min_size=(100, 100)).normalized()


def test_workspace_services_owns_one_shared_command_registry():
    services = WorkspaceServices()
    assert isinstance(services.commands, CommandRegistry)
    services.commands.register("save", lambda: "saved")
    assert services.commands.execute("save") == "saved"


def test_lifecycle_runs_guards_before_close_callbacks_once():
    calls = []
    lifecycle = WorkspaceLifecycle()
    lifecycle.add_close_guard(lambda: True)
    lifecycle.add_before_close(lambda: calls.append("persist"))
    assert lifecycle.close() is True
    assert lifecycle.close() is False
    assert calls == ["persist"]


def test_lifecycle_cancelled_close_does_not_run_persistence():
    calls = []
    lifecycle = WorkspaceLifecycle()
    lifecycle.add_close_guard(lambda: False)
    lifecycle.add_before_close(lambda: calls.append("persist"))
    assert lifecycle.close() is False
    assert lifecycle.closed is False
    assert calls == []


def test_surface_contract_supports_multiple_workspace_types():
    contract = workspace_surface_contract((
        spec("draft_tree"),
        WorkspaceSpec(
            workspace_id="classification_review",
            kind=WorkspaceKind.CLASSIFICATION_REVIEW,
            title="Review rows",
            geometry="1320x760",
            min_size=(980, 620),
            surfaces=("workspace", "command_bar", "status_bar"),
        ),
        WorkspaceSpec(
            workspace_id="sheet_mapping",
            kind=WorkspaceKind.MAPPING,
            title="Map columns",
            geometry="1220x790",
            min_size=(980, 650),
            surfaces=("workspace", "command_bar"),
        ),
    ))
    assert contract["draft_tree"][-1] == "status_bar"
    assert contract["classification_review"] == ("workspace", "command_bar", "status_bar")
    assert contract["sheet_mapping"] == ("workspace", "command_bar")


def test_surface_contract_rejects_duplicate_workspace_ids():
    with pytest.raises(ValueError):
        workspace_surface_contract((spec("draft_tree"), spec("draft_tree")))
