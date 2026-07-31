from __future__ import annotations

from enum import StrEnum


class CommandID(StrEnum):
    """Stable identifiers shared by buttons, menus, shortcuts and context menus."""

    FILE_OPEN = "file.open"
    FILE_SAVE_PROJECT = "file.save_project"
    FILE_CLOSE = "file.close"
    FILE_CLEAR_RECENT = "file.clear_recent"

    APP_ABOUT = "app.about"

    LAYOUT_TOGGLE_SIDEBAR = "layout.toggle_sidebar"
    LAYOUT_TOGGLE_COMMAND_BAR = "layout.toggle_command_bar"
    LAYOUT_TOGGLE_STATUS_BAR = "layout.toggle_status_bar"

    SHEET_FIND = "sheet.find"
    SHEET_SELECT_ALL = "sheet.select_all"
    SHEET_CLEAR_ALL = "sheet.clear_all"
    SHEET_CONFIGURE = "sheet.configure"
    SHEET_OPEN_DRAFT_TREE = "sheet.open_draft_tree"
    SHEET_USE = "sheet.use"
    SHEET_IGNORE = "sheet.ignore"
    WORKBOOK_NORMALIZE = "workbook.normalize"

    TREE_UNDO = "tree.undo"
    TREE_REDO = "tree.redo"
    TREE_SAVE_DRAFT = "tree.save_draft"
    TREE_FOCUS = "tree.focus"
    TREE_SELECT_ALL = "tree.select_all"
    TREE_INDENT = "tree.indent"
    TREE_OUTDENT = "tree.outdent"
    TREE_MOVE_UP = "tree.move_up"
    TREE_MOVE_DOWN = "tree.move_down"
    TREE_TOGGLE_TYPE = "tree.toggle_type"
    TREE_CREATE_WBS = "tree.create_wbs"
    TREE_RENAME = "tree.rename"
    TREE_DELETE = "tree.delete"
    TREE_APPROVE = "tree.approve"
    TREE_CANCEL = "tree.cancel"
