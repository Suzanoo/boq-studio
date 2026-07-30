from __future__ import annotations

from dataclasses import dataclass

from .extraction import ExtractedRow
from .values import numbering_level


@dataclass(slots=True)
class WBSState:
    wbs3: str = ""
    wbs4: str = ""
    parent_code: str = ""
    parent_counter: int = 0
    child_counter: int = 0


class WBSBuilder:
    """Stateful legacy-compatible WBS projection isolated from classification."""

    def __init__(self) -> None:
        self.state = WBSState()

    def accept_heading(self, row: ExtractedRow) -> bool:
        level = numbering_level(row.item_no) or numbering_level(row.description)
        if level == 1:
            self.state.wbs3 = row.description
            self.state.wbs4 = ""
            self.state.parent_code = row.item_no or str(self.state.parent_counter + 1)
            self.state.parent_counter += 1
            self.state.child_counter = 0
            return True
        if level is not None:
            self.state.wbs4 = row.description
            if row.item_no:
                self.state.parent_code = row.item_no
            self.state.child_counter = 0
            return True
        return False

    def build_item(self, row: ExtractedRow, sheet_name: str) -> list[object]:
        if self.state.parent_counter == 0:
            self.state.parent_counter = 1
            self.state.parent_code = "1"
        self.state.child_counter += 1
        code = (
            row.item_no
            if (numbering_level(row.item_no) or 0) >= 2
            else f"{self.state.parent_code or self.state.parent_counter}.{self.state.child_counter}"
        )
        return [
            code,
            4 if self.state.wbs4 else 3,
            "ITEM",
            "Project",
            sheet_name,
            self.state.wbs3,
            self.state.wbs4,
            row.description,
            row.unit,
            row.qty,
            row.material,
            row.labor,
            row.amount,
            sheet_name,
            row.source_row,
        ]
