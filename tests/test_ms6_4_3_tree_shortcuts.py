from boq_studio.shortcuts import shortcuts_for


def test_windows_priority_b_tree_shortcuts():
    keys = shortcuts_for("Windows")
    assert keys.move_up == ("<Alt-Up>",)
    assert keys.move_down == ("<Alt-Down>",)
    assert keys.indent == ("<Alt-Right>",)
    assert keys.outdent == ("<Alt-Left>",)


def test_macos_priority_b_tree_shortcuts_accept_option_and_alt():
    keys = shortcuts_for("Darwin")
    assert "<Option-Up>" in keys.move_up
    assert "<Alt-Up>" in keys.move_up
    assert "<Option-Down>" in keys.move_down
    assert "<Option-Right>" in keys.indent
    assert "<Option-Left>" in keys.outdent


def test_priority_b_does_not_replace_native_plain_arrow_keys():
    keys = shortcuts_for("Windows")
    all_sequences = keys.move_up + keys.move_down + keys.indent + keys.outdent
    assert "<Up>" not in all_sequences
    assert "<Down>" not in all_sequences
    assert "<Left>" not in all_sequences
    assert "<Right>" not in all_sequences


def test_priority_b_preserves_primary_select_all_contract():
    assert "<Control-a>" in shortcuts_for("Windows").select_all
    assert "<Command-a>" in shortcuts_for("Darwin").select_all

class _Widget:
    def __init__(self, widget_class: str):
        self.widget_class = widget_class

    def winfo_class(self):
        return self.widget_class


class _Event:
    def __init__(self, widget_class: str):
        self.widget = _Widget(widget_class)


def test_tree_edit_shortcut_focus_guard_without_opening_gui():
    from boq_studio.app import DraftTreeDialog

    assert DraftTreeDialog._tree_edit_shortcut_is_safe(_Event("Treeview"))
    assert not DraftTreeDialog._tree_edit_shortcut_is_safe(_Event("TEntry"))
    assert not DraftTreeDialog._tree_edit_shortcut_is_safe(_Event("Text"))


def test_move_shortcuts_route_to_existing_button_commands():
    from boq_studio.app import DraftTreeDialog

    calls = []

    class DialogStub:
        _tree_edit_shortcut_is_safe = staticmethod(DraftTreeDialog._tree_edit_shortcut_is_safe)

        def _move_up(self):
            calls.append("up")

        def _move_down(self):
            calls.append("down")

    dialog = DialogStub()
    assert DraftTreeDialog._move_up_shortcut(dialog, _Event("Treeview")) == "break"
    assert DraftTreeDialog._move_down_shortcut(dialog, _Event("Treeview")) == "break"
    assert calls == ["up", "down"]


def test_indent_shortcuts_route_to_existing_button_commands():
    from boq_studio.app import DraftTreeDialog

    calls = []

    class DialogStub:
        _tree_edit_shortcut_is_safe = staticmethod(DraftTreeDialog._tree_edit_shortcut_is_safe)

        def _indent(self):
            calls.append("indent")

        def _outdent(self):
            calls.append("outdent")

    dialog = DialogStub()
    assert DraftTreeDialog._indent_shortcut(dialog, _Event("Treeview")) == "break"
    assert DraftTreeDialog._outdent_shortcut(dialog, _Event("Treeview")) == "break"
    assert calls == ["indent", "outdent"]


def test_tree_edit_shortcuts_leave_text_editing_untouched():
    from boq_studio.app import DraftTreeDialog

    calls = []

    class DialogStub:
        _tree_edit_shortcut_is_safe = staticmethod(DraftTreeDialog._tree_edit_shortcut_is_safe)

        def _move_up(self):
            calls.append("up")

    dialog = DialogStub()
    assert DraftTreeDialog._move_up_shortcut(dialog, _Event("TEntry")) is None
    assert calls == []
