from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..models import NormalizeRequest, SheetConfig, SheetResult
from .classification import RowClassifier, RowType
from .extraction import SheetExtractor
from .wbs import WBSBuilder

ProgressCallback = Callable[[int, int, str], None]

OUTPUT_HEADERS = [
    "WBS Code", "WBS Level", "Node Type", "WBS-1", "WBS-2", "WBS-3", "WBS-4",
    "Description", "Unit", "Qty", "Material", "Labor", "Amount", "Source Sheet", "Source Row",
]


def normalize_sheet(source_ws, output_ws, cfg: SheetConfig, progress: ProgressCallback | None = None) -> SheetResult:
    result = SheetResult(sheet_name=cfg.sheet_name, status="OK")
    extractor = SheetExtractor()
    classifier = RowClassifier()
    builder = WBSBuilder()
    total_rows = max(0, source_ws.max_row - cfg.start_row + 1)

    output_ws.append(OUTPUT_HEADERS)
    for cell in output_ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for index, extracted in enumerate(extractor.extract(source_ws, cfg), start=1):
        result.source_rows += 1
        row_type = classifier.classify(extracted)
        if row_type in {RowType.EMPTY, RowType.TOTAL, RowType.NOTE}:
            result.skipped_rows += 1
        elif row_type in {RowType.WBS, RowType.UNKNOWN}:
            if builder.accept_heading(extracted):
                result.wbs_rows += 1
            result.skipped_rows += 1
        elif extracted.amount is None or extracted.amount <= 0:
            result.skipped_rows += 1
        else:
            output_ws.append(builder.build_item(extracted, cfg.sheet_name))
            result.item_rows += 1
            result.output_rows += 1

        if progress:
            progress(index, total_rows, f"{cfg.sheet_name}: row {extracted.source_row}")

    style_output_sheet(output_ws)
    if result.output_rows == 0:
        result.status = "WARNING"
        result.message = "No rows with Amount greater than zero were produced. Check the mapping."
    return result


def normalized_preview_rows(source_ws, cfg: SheetConfig, limit: int = 60) -> list[list[object]]:
    rows: list[list[object]] = []
    extractor = SheetExtractor()
    classifier = RowClassifier()
    builder = WBSBuilder()
    for extracted in extractor.extract(source_ws, cfg):
        row_type = classifier.classify(extracted)
        if row_type in {RowType.EMPTY, RowType.TOTAL, RowType.NOTE}:
            continue
        if row_type in {RowType.WBS, RowType.UNKNOWN}:
            builder.accept_heading(extracted)
            continue
        if extracted.amount is None or extracted.amount <= 0:
            continue
        rows.append(builder.build_item(extracted, cfg.sheet_name))
        if len(rows) >= limit:
            break
    return rows


def normalize_workbook(request: NormalizeRequest, progress: ProgressCallback | None = None) -> list[SheetResult]:
    source_wb = load_workbook(request.input_path, data_only=True, read_only=False)
    output_wb = Workbook()
    output_wb.remove(output_wb.active)
    results: list[SheetResult] = []
    normalized_sheet_names: list[str] = []
    total_sheets = len(request.sheets)

    for sheet_index, cfg in enumerate(request.sheets, start=1):
        if cfg.sheet_name not in source_wb.sheetnames:
            results.append(SheetResult(cfg.sheet_name, "ERROR", message="Sheet not found"))
            continue
        output_name = unique_sheet_name(output_wb, cfg.sheet_name)
        target_ws = output_wb.create_sheet(output_name)

        def sheet_progress(done: int, total: int, message: str) -> None:
            if progress:
                overall_done = (sheet_index - 1) + (done / max(1, total))
                progress(int(overall_done * 1000), total_sheets * 1000, message)

        try:
            result = normalize_sheet(source_wb[cfg.sheet_name], target_ws, cfg, sheet_progress)
            results.append(result)
            normalized_sheet_names.append(target_ws.title)
        except Exception as exc:
            output_wb.remove(target_ws)
            results.append(SheetResult(cfg.sheet_name, "ERROR", message=str(exc)))

    project_rows = add_project_sheet(output_wb, normalized_sheet_names)
    add_summary_sheet(output_wb, normalized_sheet_names)
    add_log_sheet(output_wb, request, results, project_rows)
    request.output_path.parent.mkdir(parents=True, exist_ok=True)
    output_wb.save(request.output_path)
    source_wb.close()
    output_wb.close()
    if progress:
        progress(1, 1, f"Saved: {request.output_path}")
    return results


def unique_sheet_name(wb, requested: str) -> str:
    output_name = requested[:31]
    if output_name not in wb.sheetnames:
        return output_name
    suffix = 2
    base = output_name[:27]
    while f"{base}_{suffix}" in wb.sheetnames:
        suffix += 1
    return f"{base}_{suffix}"


def style_output_sheet(ws) -> None:
    widths = {"A": 14, "B": 11, "C": 12, "D": 28, "E": 28, "F": 28, "G": 28,
              "H": 48, "I": 12, "J": 14, "K": 16, "L": 16, "M": 16, "N": 24, "O": 12}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in (10, 11, 12, 13):
        for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
            cell[0].number_format = '#,##0.00;[Red]-#,##0.00;'


def excel_sheet_ref(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def add_project_sheet(wb, normalized_sheet_names: list[str]) -> int:
    ws = wb.create_sheet("Project", 0)
    ws.append(OUTPUT_HEADERS)
    project_rows = 0
    for sheet_name in normalized_sheet_names:
        source_ws = wb[sheet_name]
        sheet_ref = excel_sheet_ref(sheet_name)
        for source_row in range(2, source_ws.max_row + 1):
            values = [source_ws.cell(source_row, column).value for column in range(1, len(OUTPUT_HEADERS) + 1)]
            if not any(value is not None and value != "" for value in values):
                continue
            project_row = ws.max_row + 1
            for column in range(1, len(OUTPUT_HEADERS) + 1):
                coordinate = source_ws.cell(source_row, column).coordinate
                ref = f"{sheet_ref}!{coordinate}"
                ws.cell(project_row, column, f'=IF({ref}="","",{ref})')
            project_rows += 1
    style_output_sheet(ws)
    return project_rows


def add_summary_sheet(wb, normalized_sheet_names: list[str]) -> None:
    ws = wb.create_sheet("SUMMARY", 0)
    ws.append(["BOQ Studio", "Amount Check"])
    ws.append(["Sheet", "Item Rows", "Amount"])
    first_detail_row = 3
    for sheet_name in normalized_sheet_names:
        source_ws = wb[sheet_name]
        last_row = max(2, source_ws.max_row)
        row_no = ws.max_row + 1
        ref = excel_sheet_ref(sheet_name)
        ws.cell(row_no, 1, sheet_name)
        ws.cell(row_no, 2, f"=MAX(COUNTA({ref}!M2:M{last_row}),0)")
        ws.cell(row_no, 3, f"=SUM({ref}!M2:M{last_row})")
    last_detail_row = ws.max_row
    grand_row = last_detail_row + 2
    ws.cell(grand_row, 1, "Grand Total")
    ws.cell(grand_row, 2, f"=SUM(B{first_detail_row}:B{last_detail_row})" if last_detail_row >= first_detail_row else 0)
    ws.cell(grand_row, 3, f"=SUM(C{first_detail_row}:C{last_detail_row})" if last_detail_row >= first_detail_row else 0)
    project_row = grand_row + 2
    project_last_row = max(2, wb["Project"].max_row)
    ws.cell(project_row, 1, "Project Sheet Total")
    ws.cell(project_row, 3, f"=SUM('Project'!M2:M{project_last_row})")
    ws.cell(project_row + 1, 1, "Difference")
    ws.cell(project_row + 1, 3, f"=C{project_row}-C{grand_row}")
    for cell in ws[1]:
        cell.font = Font(size=14, bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[grand_row]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="548235")
    ws.freeze_panes = "A3"
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 14; ws.column_dimensions["C"].width = 20
    for row in range(first_detail_row, project_row + 2):
        ws.cell(row, 3).number_format = '#,##0.00;[Red]-#,##0.00;'


def add_log_sheet(wb, request: NormalizeRequest, results: list[SheetResult], project_rows: int) -> None:
    ws = wb.create_sheet("LOG", 0)
    ws.append(["BOQ Studio R1", "Normalization Log"])
    ws.append(["Input", str(request.input_path)])
    ws.append(["Output", str(request.output_path)])
    ws.append(["Project rows", project_rows])
    ws.append([])
    ws.append(["Sheet", "Status", "Source Rows", "Output Rows", "WBS Rows", "Item Rows", "Skipped Rows", "Message"])
    for result in results:
        ws.append([result.sheet_name, result.status, result.source_rows, result.output_rows, result.wbs_rows,
                   result.item_rows, result.skipped_rows, result.message])
    for cell in ws[1]:
        cell.font = Font(size=14, bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[6]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A7"
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 18
    for col in range(3, 8):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["H"].width = 50
