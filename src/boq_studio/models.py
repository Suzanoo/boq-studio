from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class SheetConfig:
    sheet_name: str
    header_row: int = 1
    start_row: int = 2
    item_no_col: int | None = None
    description_cols: list[int] = field(default_factory=list)
    unit_col: int | None = None
    qty_col: int | None = None
    material_col: int | None = None
    labor_col: int | None = None
    amount_col: int | None = None
    max_wbs_level: int = 4
    classification_overrides: dict[int, str] = field(default_factory=dict)
    wbs_level_overrides: dict[int, int] = field(default_factory=dict)


@dataclass(slots=True)
class NormalizeRequest:
    input_path: Path
    output_path: Path
    sheets: list[SheetConfig]


@dataclass(slots=True)
class SheetResult:
    sheet_name: str
    status: str
    source_rows: int = 0
    output_rows: int = 0
    wbs_rows: int = 0
    item_rows: int = 0
    skipped_rows: int = 0
    message: str = ""
