from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .core.tree import DraftTree, TreeNodeType


class CommandBarMode(StrEnum):
    WIDE = "wide"
    COMPACT = "compact"


@dataclass(frozen=True, slots=True)
class TreeSummary:
    nodes: int
    wbs: int
    items: int
    other: int

    @property
    def primary_line(self) -> str:
        return f"{self.nodes} nodes"

    @property
    def secondary_line(self) -> str:
        parts = [f"{self.wbs} WBS", f"{self.items} Items"]
        if self.other:
            parts.append(f"{self.other} Other")
        return "  •  ".join(parts)


@dataclass(frozen=True, slots=True)
class CommandButtonSpec:
    key: str
    text: str
    shortcut: str


EDIT_COMMANDS: tuple[CommandButtonSpec, ...] = (
    CommandButtonSpec("indent", "Indent", "Alt+Right"),
    CommandButtonSpec("outdent", "Outdent", "Alt+Left"),
    CommandButtonSpec("move_up", "Move Up", "Alt+Up"),
    CommandButtonSpec("move_down", "Move Down", "Alt+Down"),
    CommandButtonSpec("toggle_type", "WBS ↔ ITEM", ""),
    CommandButtonSpec("create_wbs", "Create WBS", ""),
    CommandButtonSpec("edit", "Edit Description", "F2 / Enter"),
    CommandButtonSpec("delete", "Delete Row", "Delete"),
)


ACTION_COMMANDS: tuple[CommandButtonSpec, ...] = (
    CommandButtonSpec("save", "Save Draft", "Ctrl+S"),
    CommandButtonSpec("cancel", "Cancel", "Esc"),
    CommandButtonSpec("approve", "Use this tree", ""),
)


def command_bar_mode(width: int, *, breakpoint: int = 1080) -> CommandBarMode:
    """Return a stable responsive layout mode for the Draft Tree command bar."""
    return CommandBarMode.WIDE if width >= breakpoint else CommandBarMode.COMPACT


def editing_grid(mode: CommandBarMode, button_count: int) -> tuple[tuple[int, int], ...]:
    """Return row/column positions without hiding or dropping any command."""
    columns = button_count if mode == CommandBarMode.WIDE else 4
    return tuple((index // columns, index % columns) for index in range(button_count))


def summarize_tree(tree: DraftTree) -> TreeSummary:
    nodes = wbs = items = other = 0
    for _depth, node in tree.walk():
        nodes += 1
        if node.node_type in {TreeNodeType.WBS, TreeNodeType.VIRTUAL_WBS}:
            wbs += 1
        elif node.node_type == TreeNodeType.ITEM:
            items += 1
        else:
            other += 1
    return TreeSummary(nodes=nodes, wbs=wbs, items=items, other=other)
