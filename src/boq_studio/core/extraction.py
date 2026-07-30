from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterator

from ..models import SheetConfig
from .values import safe_number, safe_text


@dataclass(slots=True, frozen=True)
class ExtractedRow:
    source_row: int
    item_no: str
    description: str
    unit: str
    qty: float | None
    material: float | None
    labor: float | None
    amount: float | None
    raw_unit: Any
    raw_qty: Any
    is_bold: bool
    indent_level: int


class SheetExtractor:
    """Convert worksheet cells into stable, typed row records."""

    def extract(self, ws, cfg: SheetConfig) -> Iterator[ExtractedRow]:
        for row in range(cfg.start_row, ws.max_row + 1):
            yield self.extract_row(ws, row, cfg)

    def extract_row(self, ws, row: int, cfg: SheetConfig) -> ExtractedRow:
        description = " ".join(
            part for part in (safe_text(ws.cell(row, col).value) for col in cfg.description_cols) if part
        )
        item_no = safe_text(ws.cell(row, cfg.item_no_col).value) if cfg.item_no_col else ""
        raw_unit = ws.cell(row, cfg.unit_col).value if cfg.unit_col else None
        raw_qty = ws.cell(row, cfg.qty_col).value if cfg.qty_col else None
        first_description_cell = next(
            (ws.cell(row, col) for col in cfg.description_cols if safe_text(ws.cell(row, col).value)),
            None,
        )
        indent = int(first_description_cell.alignment.indent or 0) if first_description_cell else 0
        position = next(
            (index for index, col in enumerate(cfg.description_cols, start=1) if safe_text(ws.cell(row, col).value)),
            1,
        )
        return ExtractedRow(
            source_row=row,
            item_no=item_no,
            description=description,
            unit=safe_text(raw_unit),
            qty=safe_number(raw_qty),
            material=safe_number(ws.cell(row, cfg.material_col).value) if cfg.material_col else None,
            labor=safe_number(ws.cell(row, cfg.labor_col).value) if cfg.labor_col else None,
            amount=safe_number(ws.cell(row, cfg.amount_col).value) if cfg.amount_col else None,
            raw_unit=raw_unit,
            raw_qty=raw_qty,
            is_bold=any(bool(ws.cell(row, col).font.bold) for col in cfg.description_cols),
            indent_level=min(4, max(1, position + indent)),
        )
