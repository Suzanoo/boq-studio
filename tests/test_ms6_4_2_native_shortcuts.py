from boq_studio.shortcuts import is_text_input_widget, shortcuts_for


def test_windows_priority_a_shortcuts():
    keys = shortcuts_for("Windows")
    assert keys.primary == "Control"
    assert "<Control-z>" in keys.undo
    assert "<Control-y>" in keys.redo
    assert "<Control-Shift-z>" in keys.redo
    assert "<Control-s>" in keys.save
    assert "<Control-o>" in keys.open_file
    assert "<Control-f>" in keys.find
    assert "<Control-a>" in keys.select_all


def test_macos_priority_a_shortcuts():
    keys = shortcuts_for("Darwin")
    assert keys.primary == "Command"
    assert "<Command-z>" in keys.undo
    assert "<Command-Shift-z>" in keys.redo
    assert "<Command-s>" in keys.save
    assert "<Command-o>" in keys.open_file
    assert "<Command-f>" in keys.find
    assert "<Command-a>" in keys.select_all
    assert not any("Control-y" in sequence for sequence in keys.redo)


def test_delete_is_safe_inside_text_inputs():
    for widget_class in ("Entry", "TEntry", "Text", "Spinbox", "TSpinbox", "TCombobox"):
        assert is_text_input_widget(widget_class)


def test_delete_can_operate_on_tree_and_window_widgets():
    for widget_class in ("Treeview", "TButton", "TFrame", "Toplevel", None):
        assert not is_text_input_widget(widget_class)
