from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Callable, Iterable
import tkinter as tk

from .commands import CommandRegistry
from .layout import LayoutManager
from .appearance import AppearanceProfile, install_appearance


class WorkspaceKind(StrEnum):
    MAPPING = "mapping"
    CLASSIFICATION_REVIEW = "classification_review"
    DRAFT_TREE = "draft_tree"
    NORMALIZE_REVIEW = "normalize_review"


@dataclass(frozen=True, slots=True)
class WorkspaceSpec:
    """Declarative identity and window policy shared by every workspace."""

    workspace_id: str
    kind: WorkspaceKind
    title: str
    geometry: str
    min_size: tuple[int, int]
    modal: bool = True
    surfaces: tuple[str, ...] = ("workspace", "command_bar", "status_bar")

    def normalized(self) -> "WorkspaceSpec":
        width, height = self.min_size
        if not self.workspace_id.strip():
            raise ValueError("workspace_id must not be blank")
        if width < 320 or height < 240:
            raise ValueError("workspace min_size is too small")
        return WorkspaceSpec(
            workspace_id=self.workspace_id.strip(),
            kind=self.kind,
            title=self.title.strip(),
            geometry=self.geometry.strip(),
            min_size=(int(width), int(height)),
            modal=bool(self.modal),
            surfaces=tuple(dict.fromkeys(self.surfaces)),
        )


@dataclass(slots=True)
class WorkspaceServices:
    """Common services injected into a workspace instead of recreated ad hoc."""

    commands: CommandRegistry = field(default_factory=CommandRegistry)
    layout: LayoutManager | None = None
    appearance: AppearanceProfile | None = None


class WorkspaceLifecycle:
    """Small, Tk-free lifecycle coordinator used by all workspace windows."""

    def __init__(self) -> None:
        self._close_guards: list[Callable[[], bool]] = []
        self._before_close: list[Callable[[], object]] = []
        self.closed = False

    def add_close_guard(self, guard: Callable[[], bool]) -> None:
        self._close_guards.append(guard)

    def add_before_close(self, callback: Callable[[], object]) -> None:
        self._before_close.append(callback)

    def can_close(self) -> bool:
        return all(bool(guard()) for guard in tuple(self._close_guards))

    def close(self) -> bool:
        if self.closed or not self.can_close():
            return False
        for callback in tuple(self._before_close):
            callback()
        self.closed = True
        return True


class WorkspaceWindow(tk.Toplevel):
    """Shared desktop workspace host.

    It centralizes Toplevel configuration and lifecycle while leaving each
    workspace responsible for its business-specific widgets and commands.
    """

    def __init__(
        self,
        parent: tk.Misc,
        spec: WorkspaceSpec,
        *,
        services: WorkspaceServices | None = None,
    ) -> None:
        super().__init__(parent)
        self.workspace_spec = spec.normalized()
        self.workspace_services = services or WorkspaceServices()
        self.lifecycle = WorkspaceLifecycle()
        self.title(self.workspace_spec.title)
        self.geometry(self.workspace_spec.geometry)
        self.minsize(*self.workspace_spec.min_size)
        self.appearance = install_appearance(self, self.workspace_services.appearance)
        self.workspace_services.appearance = self.appearance
        self.transient(parent)
        if self.workspace_spec.modal:
            self.grab_set()

    @property
    def commands(self) -> CommandRegistry:
        return self.workspace_services.commands

    @commands.setter
    def commands(self, registry: CommandRegistry) -> None:
        self.workspace_services.commands = registry

    def request_close(self) -> bool:
        if not self.lifecycle.close():
            return False
        self.destroy()
        return True


def workspace_surface_contract(specs: Iterable[WorkspaceSpec]) -> dict[str, tuple[str, ...]]:
    """Return a deterministic contract used by tests and future shell builders."""
    result: dict[str, tuple[str, ...]] = {}
    for raw in specs:
        spec = raw.normalized()
        if spec.workspace_id in result:
            raise ValueError(f"duplicate workspace_id: {spec.workspace_id}")
        result[spec.workspace_id] = spec.surfaces
    return result
