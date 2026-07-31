from __future__ import annotations

from dataclasses import dataclass

from .extraction import ExtractedRow
from .values import numbering_level


@dataclass(slots=True)
class WBSState:
    wbs2: str = ""
    wbs3: str = ""
    wbs4: str = ""
    parent_code: str = ""
    parent_counter: int = 0
    child_counter: int = 0


def infer_wbs_level(row: ExtractedRow) -> int | None:
    """Map legacy numbering signals to output WBS levels.

    A top-level heading inside a breakdown sheet becomes WBS-3 because WBS-2
    defaults to the sheet name. Deeper numbered headings become WBS-4.
    """
    level = numbering_level(row.item_no) or numbering_level(row.description)
    if level == 1:
        return 3
    if level is not None:
        return 4
    return None


class WBSBuilder:
    """Project reviewed headings into WBS-2/WBS-3/WBS-4 state."""

    def __init__(self, sheet_name: str = "") -> None:
        self.state = WBSState(wbs2=sheet_name)

    def accept_heading(
        self,
        row: ExtractedRow,
        forced: bool = False,
        level_override: int | None = None,
    ) -> bool:
        level = level_override or infer_wbs_level(row)
        if level is None and forced:
            level = 3
        if level not in {2, 3, 4}:
            return False

        if level == 2:
            self.state.wbs2 = row.description
            self.state.wbs3 = ""
            self.state.wbs4 = ""
        elif level == 3:
            self.state.wbs3 = row.description
            self.state.wbs4 = ""
        else:
            self.state.wbs4 = row.description

        self.state.parent_counter += 1
        self.state.parent_code = row.item_no or str(self.state.parent_counter)
        self.state.child_counter = 0
        return True

    def build_item(
        self,
        row: ExtractedRow,
        sheet_name: str,
        *,
        wbs_no: str | None = None,
    ) -> list[object]:
        """Build one row for the public eleven-column export contract."""
        if self.state.parent_counter == 0:
            self.state.parent_counter = 1
            self.state.parent_code = "1"
        self.state.child_counter += 1
        code = wbs_no or (
            row.item_no
            if (numbering_level(row.item_no) or 0) >= 2
            else f"{self.state.parent_code or self.state.parent_counter}.{self.state.child_counter}"
        )
        effective_wbs2 = self.state.wbs2 or sheet_name
        return [
            code,
            "Project",
            effective_wbs2,
            self.state.wbs3,
            self.state.wbs4,
            row.description,
            row.unit,
            row.qty,
            row.material,
            row.labor,
            row.amount,
        ]
