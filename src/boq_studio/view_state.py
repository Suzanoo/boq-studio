from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Protocol

_GEOMETRY_RE = re.compile(r"^\d+x\d+(?:[+-]\d+){2}$")


@dataclass(frozen=True, slots=True)
class TreeViewState:
    """Serializable visual state for a tree workspace."""

    initialized: bool = False
    expanded_node_ids: tuple[str, ...] = ()
    selected_node_ids: tuple[str, ...] = ()
    focused_node_id: str = ""
    x_position: float = 0.0
    y_position: float = 0.0
    column_widths: tuple[tuple[str, int], ...] = ()
    window_geometry: str = ""

    def normalized(self) -> "TreeViewState":
        columns = tuple(
            (str(name), max(40, min(int(width), 2000)))
            for name, width in self.column_widths
            if str(name)
        )
        geometry = self.window_geometry if _GEOMETRY_RE.match(self.window_geometry) else ""
        return replace(
            self,
            expanded_node_ids=tuple(dict.fromkeys(map(str, self.expanded_node_ids))),
            selected_node_ids=tuple(dict.fromkeys(map(str, self.selected_node_ids))),
            focused_node_id=str(self.focused_node_id or ""),
            x_position=max(0.0, min(float(self.x_position), 1.0)),
            y_position=max(0.0, min(float(self.y_position), 1.0)),
            column_widths=columns,
            window_geometry=geometry,
        )


class ViewStateStore:
    """JSON persistence isolated from Tkinter for deterministic testing."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self, workspace_id: str) -> TreeViewState:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            raw = payload.get(workspace_id, {}) if isinstance(payload, dict) else {}
            if not isinstance(raw, dict):
                return TreeViewState()
            return TreeViewState(
                initialized=bool(raw.get("initialized", False)),
                expanded_node_ids=tuple(raw.get("expanded_node_ids", ())),
                selected_node_ids=tuple(raw.get("selected_node_ids", ())),
                focused_node_id=raw.get("focused_node_id", ""),
                x_position=raw.get("x_position", 0.0),
                y_position=raw.get("y_position", 0.0),
                column_widths=tuple(tuple(item) for item in raw.get("column_widths", ())),
                window_geometry=raw.get("window_geometry", ""),
            ).normalized()
        except (OSError, TypeError, ValueError, json.JSONDecodeError):
            return TreeViewState()

    def save(self, workspace_id: str, state: TreeViewState) -> None:
        payload: dict[str, object] = {}
        try:
            loaded = json.loads(self.path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                payload = loaded
        except (OSError, json.JSONDecodeError):
            pass
        payload[workspace_id] = asdict(state.normalized())
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


class TreeWidget(Protocol):
    def get_children(self, item: str = "") -> tuple[str, ...]: ...
    def item(self, item: str, option: str | None = None, **kwargs): ...
    def selection(self) -> tuple[str, ...]: ...
    def selection_set(self, items) -> None: ...
    def focus(self, item: str | None = None): ...
    def exists(self, item: str) -> bool: ...
    def xview(self, *args): ...
    def yview(self, *args): ...
    def column(self, column: str, option: str | None = None, **kwargs): ...


class ViewStateManager:
    """Captures and restores tree state across destructive UI refreshes."""

    def __init__(self, workspace_id: str, *, store: ViewStateStore | None = None) -> None:
        self.workspace_id = workspace_id
        self.store = store
        self.state = store.load(workspace_id) if store else TreeViewState()

    @staticmethod
    def _walk_ids(tree: TreeWidget, parent: str = ""):
        for item_id in tree.get_children(parent):
            yield item_id
            yield from ViewStateManager._walk_ids(tree, item_id)

    def capture(self, tree: TreeWidget, columns: tuple[str, ...], *, window_geometry: str = "") -> TreeViewState:
        expanded = tuple(item_id for item_id in self._walk_ids(tree) if bool(tree.item(item_id, "open")))
        xview = tree.xview()
        yview = tree.yview()
        state = TreeViewState(
            initialized=True,
            expanded_node_ids=expanded,
            selected_node_ids=tuple(tree.selection()),
            focused_node_id=str(tree.focus() or ""),
            x_position=float(xview[0]) if xview else 0.0,
            y_position=float(yview[0]) if yview else 0.0,
            column_widths=tuple((column, int(tree.column(column, "width"))) for column in columns),
            window_geometry=window_geometry,
        ).normalized()
        self.state = state
        if self.store:
            self.store.save(self.workspace_id, state)
        return state

    def restore(
        self,
        tree: TreeWidget,
        columns: tuple[str, ...],
        *,
        selected_override: tuple[str, ...] | None = None,
    ) -> None:
        state = self.state.normalized()
        if not state.initialized:
            return
        expanded = set(state.expanded_node_ids)
        for item_id in self._walk_ids(tree):
            tree.item(item_id, open=item_id in expanded)
        widths = dict(state.column_widths)
        for column in columns:
            if column in widths:
                tree.column(column, width=widths[column])
        selected = selected_override if selected_override is not None else state.selected_node_ids
        visible = tuple(item_id for item_id in selected if tree.exists(item_id))
        tree.selection_set(visible)
        focus = visible[0] if selected_override and visible else state.focused_node_id
        if focus and tree.exists(focus):
            tree.focus(focus)
        tree.xview("moveto", state.x_position)
        tree.yview("moveto", state.y_position)

    def restore_window_geometry(self, window) -> None:
        if self.state.window_geometry:
            window.geometry(self.state.window_geometry)
