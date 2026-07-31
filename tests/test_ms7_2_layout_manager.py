import json
from dataclasses import replace

import pytest

from boq_studio.layout import LayoutManager, LayoutState, LayoutStateStore
from boq_studio.native_os import layout_state_path


class FakeRegion:
    def __init__(self) -> None:
        self.visible = None
        self.show_calls = 0
        self.hide_calls = 0

    def show(self) -> None:
        self.visible = True
        self.show_calls += 1

    def hide(self) -> None:
        self.visible = False
        self.hide_calls += 1

    def is_visible(self) -> bool:
        return bool(self.visible)


def test_layout_manager_applies_saved_state_when_region_is_registered(tmp_path):
    store = LayoutStateStore(tmp_path / "layout.json")
    store.save("draft_tree", LayoutState(sidebar_visible=False))
    manager = LayoutManager("draft_tree", state_store=store)
    sidebar = FakeRegion()

    manager.register_region("sidebar", sidebar)

    assert sidebar.visible is False
    assert sidebar.hide_calls == 1


def test_toggle_updates_region_state_and_persists_it(tmp_path):
    store = LayoutStateStore(tmp_path / "layout.json")
    manager = LayoutManager("draft_tree", state_store=store)
    sidebar = FakeRegion()
    manager.register_region("sidebar", sidebar)

    assert manager.toggle("sidebar") is False
    assert sidebar.visible is False

    reloaded = LayoutManager("draft_tree", state_store=store)
    assert reloaded.state.sidebar_visible is False


def test_layout_state_is_kept_separately_for_each_workspace(tmp_path):
    store = LayoutStateStore(tmp_path / "layout.json")
    store.save("draft_tree", LayoutState(sidebar_visible=False))
    store.save("classification", LayoutState(command_bar_visible=False))

    draft = store.load("draft_tree")
    classification = store.load("classification")

    assert draft.sidebar_visible is False
    assert draft.command_bar_visible is True
    assert classification.sidebar_visible is True
    assert classification.command_bar_visible is False


def test_sidebar_width_is_clamped_to_safe_desktop_limits():
    assert LayoutState(sidebar_width=20).normalized().sidebar_width == 180
    assert LayoutState(sidebar_width=900).normalized().sidebar_width == 640
    assert LayoutState(sidebar_width=320).normalized().sidebar_width == 320


def test_setting_sidebar_width_persists_normalized_value(tmp_path):
    store = LayoutStateStore(tmp_path / "layout.json")
    manager = LayoutManager("draft_tree", state_store=store)

    assert manager.set_sidebar_width(999) == 640
    assert store.load("draft_tree").sidebar_width == 640


def test_unknown_region_is_rejected():
    manager = LayoutManager("draft_tree")
    with pytest.raises(KeyError):
        manager.toggle("unknown")


def test_corrupt_layout_file_falls_back_to_default(tmp_path):
    path = tmp_path / "layout.json"
    path.write_text("not json", encoding="utf-8")
    state = LayoutStateStore(path).load("draft_tree", LayoutState(sidebar_visible=False))
    assert state.sidebar_visible is False


def test_layout_store_preserves_other_workspace_payload(tmp_path):
    path = tmp_path / "layout.json"
    path.write_text(json.dumps({"other": {"sidebar_visible": False}}), encoding="utf-8")
    store = LayoutStateStore(path)
    store.save("draft_tree", LayoutState(status_bar_visible=False))
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["other"]["sidebar_visible"] is False
    assert payload["draft_tree"]["status_bar_visible"] is False


def test_layout_path_uses_shared_application_config_directory(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path / "AppData"))
    assert layout_state_path("Windows", tmp_path) == tmp_path / "AppData" / "BOQ Studio" / "layout.json"


def test_draft_tree_source_uses_layout_manager_and_shared_layout_commands():
    from pathlib import Path

    source = Path("src/boq_studio/app.py").read_text(encoding="utf-8")
    assert 'LayoutManager(\n            "draft_tree"' in source
    assert 'register(CommandID.LAYOUT_TOGGLE_SIDEBAR, self._toggle_sidebar)' in source
    assert 'self.layout.register_region("sidebar", GridRegion(self.sidebar))' in source
    assert 'self.commands = CommandRegistry()' in source
    assert 'self._register_commands()' in source
