"""Backward-compatible facade for the BOQ normalization pipeline.

R1 moved implementation into small UI-independent core modules. Existing GUI,
tests, scripts, and external imports can continue using this module unchanged.
"""
from __future__ import annotations

from .core.exporter import (
    OUTPUT_HEADERS,
    ProgressCallback,
    add_log_sheet,
    add_project_sheet,
    add_summary_sheet,
    excel_sheet_ref,
    normalize_sheet,
    normalize_workbook,
    normalized_preview_rows,
    style_output_sheet,
)
from .core.values import is_blank_value, numbering_level, safe_number, safe_text

__all__ = [
    "OUTPUT_HEADERS", "ProgressCallback", "add_log_sheet", "add_project_sheet", "add_summary_sheet",
    "excel_sheet_ref", "is_blank_value", "normalize_sheet", "normalize_workbook",
    "normalized_preview_rows", "numbering_level", "safe_number", "safe_text", "style_output_sheet",
]
