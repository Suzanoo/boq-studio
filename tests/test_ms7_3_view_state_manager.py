import json
from pathlib import Path

from boq_studio.native_os import view_state_path
from boq_studio.view_state import TreeViewState, ViewStateManager, ViewStateStore


class FakeTree:
    def __init__(self):
        self.children = {"": ["root"], "root": ["floor", "wall", "ceiling"], "floor": ["f1"], "wall": [], "ceiling": ["c1"], "f1": [], "c1": []}
        self.open = {key: False for key in self.children}
        self.selected = ()
        self.focused = ""
        self.x = (0.0, 1.0)
        self.y = (0.0, 1.0)
        self.widths = {"#0": 650, "type": 110, "amount": 130}

    def get_children(self, item=""):
        return tuple(self.children.get(item, ()))

    def item(self, item, option=None, **kwargs):
        if "open" in kwargs:
            self.open[item] = bool(kwargs["open"])
        if option == "open":
            return self.open[item]
        return {"open": self.open[item]}

    def selection(self):
        return self.selected

    def selection_set(self, items):
        self.selected = tuple(items) if not isinstance(items, str) else (items,)

    def focus(self, item=None):
        if item is not None:
            self.focused = item
        return self.focused

    def exists(self, item):
        return item in self.children

    def xview(self, *args):
        if args:
            self.x = (float(args[1]), 1.0)
        return self.x

    def yview(self, *args):
        if args:
            self.y = (float(args[1]), 1.0)
        return self.y

    def column(self, column, option=None, **kwargs):
        if "width" in kwargs:
            self.widths[column] = kwargs["width"]
        if option == "width":
            return self.widths[column]
        return {"width": self.widths[column]}


def test_capture_and_restore_preserves_collapsed_branches_selection_focus_and_scroll():
    tree = FakeTree()
    tree.open.update({"root": True, "floor": False, "wall": False, "ceiling": True})
    tree.selected = ("c1",)
    tree.focused = "c1"
    tree.x = (0.25, 0.75)
    tree.y = (0.4, 0.8)
    manager = ViewStateManager("draft")
    manager.capture(tree, ("#0", "type", "amount"), window_geometry="1050x760+50+40")

    tree.open.update({key: True for key in tree.open})
    tree.selected = ()
    tree.focused = ""
    tree.x = (0.0, 1.0)
    tree.y = (0.0, 1.0)
    tree.widths["#0"] = 300
    manager.restore(tree, ("#0", "type", "amount"))

    assert tree.open["floor"] is False
    assert tree.open["wall"] is False
    assert tree.open["ceiling"] is True
    assert tree.selected == ("c1",)
    assert tree.focused == "c1"
    assert tree.x[0] == 0.25
    assert tree.y[0] == 0.4
    assert tree.widths["#0"] == 650


def test_selection_override_keeps_edit_target_without_expanding_other_branches():
    tree = FakeTree()
    manager = ViewStateManager("draft")
    manager.state = TreeViewState(initialized=True, expanded_node_ids=("root", "ceiling"), selected_node_ids=("floor",))
    manager.restore(tree, ("#0", "type", "amount"), selected_override=("c1",))
    assert tree.selected == ("c1",)
    assert tree.focused == "c1"
    assert tree.open["floor"] is False
    assert tree.open["ceiling"] is True


def test_deleted_nodes_are_ignored_during_restore():
    tree = FakeTree()
    manager = ViewStateManager("draft")
    manager.state = TreeViewState(initialized=True, expanded_node_ids=("missing", "root"), selected_node_ids=("missing",), focused_node_id="missing")
    manager.restore(tree, ("#0", "type", "amount"))
    assert tree.selected == ()
    assert tree.focused == ""
    assert tree.open["root"] is True


def test_store_round_trip_and_workspace_isolation(tmp_path):
    store = ViewStateStore(tmp_path / "view_state.json")
    first = TreeViewState(initialized=True, expanded_node_ids=("a",), window_geometry="900x700+10+20")
    second = TreeViewState(initialized=True, expanded_node_ids=("b",))
    store.save("first", first)
    store.save("second", second)
    assert store.load("first").expanded_node_ids == ("a",)
    assert store.load("first").window_geometry == "900x700+10+20"
    assert store.load("second").expanded_node_ids == ("b",)


def test_invalid_values_are_normalized():
    state = TreeViewState(
        initialized=True,
        x_position=-2,
        y_position=4,
        column_widths=(("#0", 5), ("amount", 9000)),
        window_geometry="bad",
    ).normalized()
    assert state.x_position == 0.0
    assert state.y_position == 1.0
    assert state.column_widths == (("#0", 40), ("amount", 2000))
    assert state.window_geometry == ""


def test_corrupt_store_falls_back_to_empty_state(tmp_path):
    path = tmp_path / "view_state.json"
    path.write_text("broken", encoding="utf-8")
    assert ViewStateStore(path).load("draft") == TreeViewState()


def test_view_state_path_uses_application_config_directory(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path / "AppData"))
    assert view_state_path("Windows", tmp_path) == tmp_path / "AppData" / "BOQ Studio" / "view_state.json"


def test_draft_tree_source_routes_refresh_through_view_state_manager():
    source = Path("src/boq_studio/app.py").read_text(encoding="utf-8")
    assert "self.view_state = ViewStateManager(" in source
    assert "self.view_state.capture(self.tree, columns" in source
    assert "self.view_state.restore(self.tree, columns" in source
    assert "self.view_state.restore_window_geometry(self)" in source
