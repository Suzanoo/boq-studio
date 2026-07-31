from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .commands import CommandID


class Surface(StrEnum):
    MENU = "menu"
    TOOLBAR = "toolbar"
    CONTEXT = "context"


@dataclass(frozen=True, slots=True)
class CommandPresentation:
    command_id: CommandID
    label: str
    shortcut: str = ""
    tooltip: str = ""


TREE_FILE_MENU: tuple[CommandPresentation, ...] = (
    CommandPresentation(CommandID.TREE_SAVE_DRAFT, "Save Draft", "Ctrl+S", "Save the current Draft WBS Tree."),
    CommandPresentation(CommandID.TREE_APPROVE, "Use this tree", "", "Approve this draft and return it to the workflow."),
    CommandPresentation(CommandID.TREE_CANCEL, "Close", "Esc", "Close the workspace without approving the tree."),
)

TREE_EDIT_MENU: tuple[CommandPresentation, ...] = (
    CommandPresentation(CommandID.TREE_UNDO, "Undo", "Ctrl+Z", "Undo the most recent tree edit."),
    CommandPresentation(CommandID.TREE_REDO, "Redo", "Ctrl+Y", "Redo the most recently undone tree edit."),
    CommandPresentation(CommandID.TREE_SELECT_ALL, "Select All", "Ctrl+A", "Select every editable tree node."),
    CommandPresentation(CommandID.TREE_RENAME, "Edit Description", "F2", "Edit the selected node description."),
    CommandPresentation(CommandID.TREE_DELETE, "Delete Row", "Delete", "Delete the selected row or rows."),
)

TREE_MENU: tuple[CommandPresentation, ...] = (
    CommandPresentation(CommandID.TREE_INDENT, "Indent", "Alt+Right", "Move the selected node one WBS level deeper."),
    CommandPresentation(CommandID.TREE_OUTDENT, "Outdent", "Alt+Left", "Move the selected node one WBS level higher."),
    CommandPresentation(CommandID.TREE_MOVE_UP, "Move Up", "Alt+Up", "Move selected rows upward within their parent."),
    CommandPresentation(CommandID.TREE_MOVE_DOWN, "Move Down", "Alt+Down", "Move selected rows downward within their parent."),
    CommandPresentation(CommandID.TREE_TOGGLE_TYPE, "WBS ↔ ITEM", "", "Convert the selected node between WBS and ITEM."),
    CommandPresentation(CommandID.TREE_CREATE_WBS, "Create WBS", "", "Create a new WBS heading near the selection."),
    CommandPresentation(CommandID.TREE_EXPAND_SELECTED, "Expand Selected", "Right", "Expand the selected tree branch."),
    CommandPresentation(CommandID.TREE_COLLAPSE_SELECTED, "Collapse Selected", "Left", "Collapse the selected tree branch."),
    CommandPresentation(CommandID.TREE_EXPAND_ALL, "Expand All", "", "Expand all WBS branches."),
    CommandPresentation(CommandID.TREE_COLLAPSE_ALL, "Collapse All", "", "Collapse all WBS branches."),
)

TREE_VIEW_MENU: tuple[CommandPresentation, ...] = (
    CommandPresentation(CommandID.LAYOUT_TOGGLE_SIDEBAR, "Node Details", "", "Show or hide the Node Details sidebar."),
    CommandPresentation(CommandID.LAYOUT_TOGGLE_COMMAND_BAR, "Command Bar", "", "Show or hide the bottom editing command bar."),
    CommandPresentation(CommandID.LAYOUT_TOGGLE_STATUS_BAR, "Status Bar", "", "Show or hide the workspace status bar."),
)

TREE_TOOLBAR: tuple[CommandPresentation, ...] = (
    TREE_FILE_MENU[0],
    TREE_EDIT_MENU[0], TREE_EDIT_MENU[1],
    TREE_MENU[0], TREE_MENU[1], TREE_MENU[2], TREE_MENU[3],
    TREE_MENU[5], TREE_EDIT_MENU[3], TREE_EDIT_MENU[4],
)

TREE_CONTEXT_MENU: tuple[CommandPresentation, ...] = (
    TREE_EDIT_MENU[3],
    TREE_MENU[5],
    TREE_MENU[0], TREE_MENU[1], TREE_MENU[2], TREE_MENU[3],
    TREE_MENU[6], TREE_MENU[7],
    TREE_MENU[4],
    TREE_EDIT_MENU[4],
)


def normalized_shortcut(shortcut: str, *, system_name: str) -> str:
    if system_name == "Darwin":
        return shortcut.replace("Ctrl+", "⌘").replace("Alt+", "Option+")
    return shortcut


def tooltip_text(spec: CommandPresentation, *, system_name: str) -> str:
    shortcut = normalized_shortcut(spec.shortcut, system_name=system_name)
    return spec.tooltip if not shortcut else f"{spec.tooltip}\nShortcut: {shortcut}"
