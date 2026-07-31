from __future__ import annotations

import json
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Protocol


@dataclass(frozen=True, slots=True)
class LayoutState:
    """Persistent visibility and sizing preferences for one workspace."""

    sidebar_visible: bool = True
    command_bar_visible: bool = True
    status_bar_visible: bool = True
    toolbar_visible: bool = True
    sidebar_width: int = 280

    def normalized(self) -> "LayoutState":
        return replace(self, sidebar_width=max(180, min(int(self.sidebar_width), 640)))


class LayoutStateStore:
    """JSON persistence for workspace layout state.

    The store is intentionally independent from Tkinter so layout behavior can be
    tested without opening a desktop window.
    """

    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self, workspace_id: str, default: LayoutState | None = None) -> LayoutState:
        fallback = default or LayoutState()
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            raw = payload.get(workspace_id, {}) if isinstance(payload, dict) else {}
            if not isinstance(raw, dict):
                return fallback.normalized()
            values = asdict(fallback)
            values.update({key: raw[key] for key in values if key in raw})
            return LayoutState(**values).normalized()
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            return fallback.normalized()

    def save(self, workspace_id: str, state: LayoutState) -> None:
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


class VisibilityRegion(Protocol):
    def show(self) -> None: ...
    def hide(self) -> None: ...
    def is_visible(self) -> bool: ...


class LayoutManager:
    """Owns workspace region visibility and persistent layout preferences."""

    def __init__(
        self,
        workspace_id: str,
        *,
        state_store: LayoutStateStore | None = None,
        initial_state: LayoutState | None = None,
    ) -> None:
        self.workspace_id = workspace_id
        self.state_store = state_store
        self.state = (
            state_store.load(workspace_id, initial_state)
            if state_store is not None
            else (initial_state or LayoutState()).normalized()
        )
        self._regions: dict[str, VisibilityRegion] = {}

    def register_region(self, name: str, region: VisibilityRegion) -> None:
        self._regions[name] = region
        self._apply_region(name)

    def set_visible(self, name: str, visible: bool) -> bool:
        field = f"{name}_visible"
        if not hasattr(self.state, field):
            raise KeyError(f"Unknown layout region: {name}")
        self.state = replace(self.state, **{field: bool(visible)})
        self._apply_region(name)
        self.persist()
        return bool(visible)

    def toggle(self, name: str) -> bool:
        field = f"{name}_visible"
        if not hasattr(self.state, field):
            raise KeyError(f"Unknown layout region: {name}")
        return self.set_visible(name, not bool(getattr(self.state, field)))

    def set_sidebar_width(self, width: int) -> int:
        self.state = replace(self.state, sidebar_width=width).normalized()
        self.persist()
        return self.state.sidebar_width

    def persist(self) -> None:
        if self.state_store is not None:
            self.state_store.save(self.workspace_id, self.state)

    def _apply_region(self, name: str) -> None:
        region = self._regions.get(name)
        if region is None:
            return
        visible = bool(getattr(self.state, f"{name}_visible"))
        region.show() if visible else region.hide()


class GridRegion:
    """Tk-compatible region adapter using grid/grid_remove without importing Tk."""

    def __init__(self, widget) -> None:
        self.widget = widget

    def show(self) -> None:
        self.widget.grid()

    def hide(self) -> None:
        self.widget.grid_remove()

    def is_visible(self) -> bool:
        return bool(self.widget.winfo_ismapped())


class PackRegion:
    """Tk-compatible region adapter preserving the original pack configuration."""

    def __init__(self, widget, **pack_options) -> None:
        self.widget = widget
        self.pack_options = pack_options

    def show(self) -> None:
        self.widget.pack(**self.pack_options)

    def hide(self) -> None:
        self.widget.pack_forget()

    def is_visible(self) -> bool:
        return bool(self.widget.winfo_ismapped())
