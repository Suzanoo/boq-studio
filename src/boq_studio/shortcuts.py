from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ShortcutSet:
    primary: str
    undo: tuple[str, ...]
    redo: tuple[str, ...]
    save: tuple[str, ...]
    open_file: tuple[str, ...]
    find: tuple[str, ...]
    select_all: tuple[str, ...]
    move_up: tuple[str, ...]
    move_down: tuple[str, ...]
    indent: tuple[str, ...]
    outdent: tuple[str, ...]


def shortcuts_for(system_name: str) -> ShortcutSet:
    """Return Tk event sequences for the target operating system."""
    is_mac = system_name.casefold() == "darwin"
    primary = "Command" if is_mac else "Control"
    redo = (f"<{primary}-Shift-z>", f"<{primary}-Shift-Z>")
    if not is_mac:
        redo += ("<Control-y>", "<Control-Y>")
    option = "Option" if is_mac else "Alt"
    alternate_option = ("<Alt-Up>",) if is_mac else ()
    alternate_option_down = ("<Alt-Down>",) if is_mac else ()
    alternate_option_right = ("<Alt-Right>",) if is_mac else ()
    alternate_option_left = ("<Alt-Left>",) if is_mac else ()
    return ShortcutSet(
        primary=primary,
        undo=(f"<{primary}-z>", f"<{primary}-Z>"),
        redo=redo,
        save=(f"<{primary}-s>", f"<{primary}-S>"),
        open_file=(f"<{primary}-o>", f"<{primary}-O>"),
        find=(f"<{primary}-f>", f"<{primary}-F>"),
        select_all=(f"<{primary}-a>", f"<{primary}-A>"),
        move_up=(f"<{option}-Up>",) + alternate_option,
        move_down=(f"<{option}-Down>",) + alternate_option_down,
        indent=(f"<{option}-Right>",) + alternate_option_right,
        outdent=(f"<{option}-Left>",) + alternate_option_left,
    )


def is_text_input_widget(widget_class: str | None) -> bool:
    """True when Delete/Copy/Paste should remain native text-edit commands."""
    return (widget_class or "").casefold() in {
        "entry",
        "tentry",
        "text",
        "spinbox",
        "tspinbox",
        "combobox",
        "tcombobox",
    }
