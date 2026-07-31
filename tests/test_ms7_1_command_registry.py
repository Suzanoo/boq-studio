import pytest

from boq_studio.commands import (
    CommandDisabledError,
    CommandID,
    CommandNotRegisteredError,
    CommandRegistry,
)


def test_registry_executes_one_handler_for_every_ui_source():
    calls = []
    registry = CommandRegistry()
    registry.register(CommandID.TREE_DELETE, lambda: calls.append("delete"))

    button_callback = registry.callback(CommandID.TREE_DELETE)
    menu_callback = registry.callback(CommandID.TREE_DELETE)
    context_callback = registry.callback(CommandID.TREE_DELETE)
    shortcut_callback = registry.event_callback(CommandID.TREE_DELETE)

    button_callback()
    menu_callback()
    context_callback()
    assert shortcut_callback(object()) == "break"
    assert calls == ["delete", "delete", "delete", "delete"]


def test_registry_rejects_duplicate_ids_by_default():
    registry = CommandRegistry()
    registry.register(CommandID.FILE_OPEN, lambda: None)
    with pytest.raises(ValueError):
        registry.register(CommandID.FILE_OPEN, lambda: None)


def test_registry_supports_explicit_handler_replacement():
    calls = []
    registry = CommandRegistry()
    registry.register(CommandID.FILE_OPEN, lambda: calls.append("old"))
    registry.register(CommandID.FILE_OPEN, lambda: calls.append("new"), replace=True)
    registry.execute(CommandID.FILE_OPEN)
    assert calls == ["new"]


def test_disabled_command_does_not_execute():
    calls = []
    registry = CommandRegistry()
    registry.register(
        CommandID.WORKBOOK_NORMALIZE,
        lambda: calls.append("run"),
        can_execute=lambda: False,
    )
    assert not registry.is_enabled(CommandID.WORKBOOK_NORMALIZE)
    with pytest.raises(CommandDisabledError):
        registry.execute(CommandID.WORKBOOK_NORMALIZE)
    assert calls == []


def test_unknown_command_has_clear_error():
    registry = CommandRegistry()
    with pytest.raises(CommandNotRegisteredError):
        registry.execute("missing.command")


def test_event_callback_can_pass_event_to_focus_guard():
    seen = []
    registry = CommandRegistry()
    registry.register(CommandID.TREE_SELECT_ALL, lambda event=None: seen.append(event) or None)
    event = object()
    callback = registry.event_callback(CommandID.TREE_SELECT_ALL, pass_event=True)
    assert callback(event) == "break"
    assert seen == [event]


def test_command_ids_are_stable_namespaced_values():
    assert CommandID.FILE_OPEN == "file.open"
    assert CommandID.SHEET_CONFIGURE == "sheet.configure"
    assert CommandID.TREE_DELETE == "tree.delete"
