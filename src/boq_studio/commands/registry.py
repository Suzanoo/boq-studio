from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Hashable

CommandKey = Hashable
CommandHandler = Callable[..., object]
CommandPredicate = Callable[[], bool]


class CommandError(RuntimeError):
    """Base error raised by the command layer."""


class CommandNotRegisteredError(CommandError, KeyError):
    """Raised when a caller executes an unknown command ID."""


class CommandDisabledError(CommandError):
    """Raised when a registered command is currently disabled."""


@dataclass(frozen=True)
class Command:
    command_id: CommandKey
    handler: CommandHandler
    can_execute: CommandPredicate = lambda: True
    description: str = ""

    @property
    def enabled(self) -> bool:
        return bool(self.can_execute())


class CommandRegistry:
    """Single routing point for all UI command sources.

    The registry deliberately contains no Tkinter dependency, which keeps command
    routing testable without opening a GUI.
    """

    def __init__(self) -> None:
        self._commands: dict[CommandKey, Command] = {}

    def register(
        self,
        command_id: CommandKey,
        handler: CommandHandler,
        *,
        can_execute: CommandPredicate | None = None,
        description: str = "",
        replace: bool = False,
    ) -> Command:
        if command_id in self._commands and not replace:
            raise ValueError(f"Command already registered: {command_id}")
        command = Command(command_id, handler, can_execute or (lambda: True), description)
        self._commands[command_id] = command
        return command

    def unregister(self, command_id: CommandKey) -> None:
        self._commands.pop(command_id, None)

    def get(self, command_id: CommandKey) -> Command:
        try:
            return self._commands[command_id]
        except KeyError as exc:
            raise CommandNotRegisteredError(command_id) from exc

    def contains(self, command_id: CommandKey) -> bool:
        return command_id in self._commands

    def is_enabled(self, command_id: CommandKey) -> bool:
        return self.get(command_id).enabled

    def execute(self, command_id: CommandKey, *args, **kwargs):
        command = self.get(command_id)
        if not command.enabled:
            raise CommandDisabledError(f"Command is disabled: {command_id}")
        return command.handler(*args, **kwargs)

    def callback(self, command_id: CommandKey, *args, **kwargs) -> Callable[[], object]:
        """Return a zero-argument callback suitable for Tk ``command=``."""
        return lambda: self.execute(command_id, *args, **kwargs)

    def event_callback(
        self,
        command_id: CommandKey,
        *,
        pass_event: bool = False,
        break_result: bool = True,
    ) -> Callable[[object], object]:
        """Return a callback suitable for Tk ``bind``.

        ``pass_event`` is opt-in because most application commands should not know
        about Tk events. Event-aware safety guards can still be registered normally.
        """

        def invoke(event=None):
            result = self.execute(command_id, event) if pass_event else self.execute(command_id)
            if result is not None:
                return result
            return "break" if break_result else None

        return invoke

    def ids(self) -> tuple[CommandKey, ...]:
        return tuple(self._commands)
