from boq_studio.commands import CommandID
from boq_studio.desktop_ui import (
    TREE_CONTEXT_MENU,
    TREE_EDIT_MENU,
    TREE_FILE_MENU,
    TREE_MENU,
    TREE_TOOLBAR,
    TREE_VIEW_MENU,
    normalized_shortcut,
    tooltip_text,
)


def ids(specs):
    return tuple(spec.command_id for spec in specs)


def test_menu_surfaces_use_registered_command_ids():
    all_specs = TREE_FILE_MENU + TREE_EDIT_MENU + TREE_MENU + TREE_VIEW_MENU
    assert all(isinstance(spec.command_id, CommandID) for spec in all_specs)
    assert len(ids(all_specs)) == len(set(ids(all_specs)))


def test_file_menu_contains_save_approve_and_close():
    assert ids(TREE_FILE_MENU) == (
        CommandID.TREE_SAVE_DRAFT,
        CommandID.TREE_APPROVE,
        CommandID.TREE_CANCEL,
    )


def test_tree_menu_contains_existing_edit_commands_and_view_commands():
    tree_ids = set(ids(TREE_MENU))
    assert {
        CommandID.TREE_INDENT,
        CommandID.TREE_OUTDENT,
        CommandID.TREE_MOVE_UP,
        CommandID.TREE_MOVE_DOWN,
        CommandID.TREE_TOGGLE_TYPE,
        CommandID.TREE_CREATE_WBS,
    } <= tree_ids
    assert set(ids(TREE_VIEW_MENU)) == {
        CommandID.LAYOUT_TOGGLE_SIDEBAR,
        CommandID.LAYOUT_TOGGLE_COMMAND_BAR,
        CommandID.LAYOUT_TOGGLE_STATUS_BAR,
    }


def test_toolbar_is_compact_subset_of_menu_commands():
    menu_ids = set(ids(TREE_FILE_MENU + TREE_EDIT_MENU + TREE_MENU + TREE_VIEW_MENU))
    assert set(ids(TREE_TOOLBAR)) <= menu_ids
    assert len(TREE_TOOLBAR) == len(set(ids(TREE_TOOLBAR)))


def test_context_menu_has_no_new_business_logic_commands():
    allowed = set(ids(TREE_EDIT_MENU + TREE_MENU))
    assert set(ids(TREE_CONTEXT_MENU)) <= allowed
    assert CommandID.TREE_DELETE in ids(TREE_CONTEXT_MENU)
    assert CommandID.TREE_RENAME in ids(TREE_CONTEXT_MENU)


def test_expand_and_collapse_commands_are_complete_pairs():
    tree_ids = set(ids(TREE_MENU))
    assert {CommandID.TREE_EXPAND_SELECTED, CommandID.TREE_COLLAPSE_SELECTED} <= tree_ids
    assert {CommandID.TREE_EXPAND_ALL, CommandID.TREE_COLLAPSE_ALL} <= tree_ids


def test_shortcuts_use_native_macos_labels_without_changing_windows():
    assert normalized_shortcut("Ctrl+S", system_name="Windows") == "Ctrl+S"
    assert normalized_shortcut("Alt+Right", system_name="Windows") == "Alt+Right"
    assert normalized_shortcut("Ctrl+S", system_name="Darwin") == "⌘S"
    assert normalized_shortcut("Alt+Right", system_name="Darwin") == "Option+Right"


def test_tooltip_includes_description_and_native_shortcut():
    save = TREE_FILE_MENU[0]
    assert tooltip_text(save, system_name="Windows").endswith("Shortcut: Ctrl+S")
    assert tooltip_text(save, system_name="Darwin").endswith("Shortcut: ⌘S")
    approve = TREE_FILE_MENU[1]
    assert "Shortcut:" not in tooltip_text(approve, system_name="Windows")
