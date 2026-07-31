from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Iterable


class SheetState(str, Enum):
    NEW = "NEW"
    MAPPED = "MAPPED"
    TREE_DRAFT = "TREE DRAFT"
    TREE_READY = "TREE READY"
    EXPORTED = "EXPORTED"


@dataclass(frozen=True, slots=True)
class WorkflowStatus:
    selected_count: int
    new_count: int
    mapped_count: int
    tree_draft_count: int
    tree_ready_count: int
    exported_count: int

    @property
    def can_export(self) -> bool:
        return self.selected_count > 0 and self.new_count == 0 and self.mapped_count == 0 and self.tree_draft_count == 0

    @property
    def summary(self) -> str:
        if self.selected_count == 0:
            return "No breakdown sheets selected"
        if self.new_count:
            return f"{self.selected_count} selected · {self.new_count} need mapping"
        if self.mapped_count:
            return f"{self.selected_count} selected · {self.mapped_count} need Draft WBS Tree"
        if self.tree_draft_count:
            return f"{self.selected_count} selected · {self.tree_draft_count} tree draft(s) not approved"
        if self.exported_count == self.selected_count:
            return f"{self.selected_count} selected · Export completed"
        return f"{self.selected_count} selected · Ready to normalize"


def sheet_state(*, has_mapping: bool, has_tree: bool, tree_ready: bool = False, exported: bool = False) -> SheetState:
    if exported and has_tree:
        return SheetState.EXPORTED
    if has_tree and tree_ready:
        return SheetState.TREE_READY
    if has_tree:
        return SheetState.TREE_DRAFT
    if has_mapping:
        return SheetState.MAPPED
    return SheetState.NEW


def workflow_status(states: Iterable[SheetState]) -> WorkflowStatus:
    values = list(states)
    return WorkflowStatus(
        selected_count=len(values),
        new_count=sum(value is SheetState.NEW for value in values),
        mapped_count=sum(value is SheetState.MAPPED for value in values),
        tree_draft_count=sum(value is SheetState.TREE_DRAFT for value in values),
        tree_ready_count=sum(value is SheetState.TREE_READY for value in values),
        exported_count=sum(value is SheetState.EXPORTED for value in values),
    )
