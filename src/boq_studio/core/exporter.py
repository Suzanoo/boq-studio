from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

from ..models import NormalizeRequest, SheetConfig, SheetResult
from .classification import RowClassifier, RowType, resolve_row_type
from .extraction import SheetExtractor
from .wbs import WBSBuilder
from .tree import DraftTree, DraftTreeNode, TreeNodeType

ProgressCallback = Callable[[int, int, str], None]

# Public workbook contract shared by BOQ Studio and Progress Studio.
# Internal identifiers (Node ID, parent ID, source row, workflow state, etc.)
# remain in the hidden project/session JSON and are never exposed here.
OUTPUT_HEADERS = [
    "WBS No.",
    "WBS-1",
    "WBS-2",
    "WBS-3",
    "WBS-4",
    "Description",
    "Unit",
    "Qty",
    "Material",
    "Labor",
    "Amount",
]


def _style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def normalize_sheet(
    source_ws,
    output_ws,
    cfg: SheetConfig,
    progress: ProgressCallback | None = None,
    *,
    project_number: int = 1,
) -> SheetResult:
    """Legacy non-tree export.

    Guided Workflow normally exports an approved Draft Tree. This fallback keeps
    older scripts working while still honoring the nine-column workbook
    contract and project-wide numbering prefix.
    """

    result = SheetResult(sheet_name=cfg.sheet_name, status="OK")
    extractor = SheetExtractor()
    classifier = RowClassifier()
    builder = WBSBuilder(cfg.sheet_name)
    total_rows = max(0, source_ws.max_row - cfg.start_row + 1)

    output_ws.append(OUTPUT_HEADERS)
    _style_header(output_ws)

    extracted_rows = list(extractor.extract(source_ws, cfg))
    decisions = classifier.classify_rows(extracted_rows)
    item_counter = 0
    for index, (extracted, decision) in enumerate(zip(extracted_rows, decisions, strict=True), start=1):
        result.source_rows += 1
        row_type = resolve_row_type(decision, extracted.source_row, cfg.classification_overrides)
        if row_type in {RowType.EMPTY, RowType.TOTAL, RowType.NOTE, RowType.DETAIL, RowType.IGNORE}:
            result.skipped_rows += 1
        elif row_type in {RowType.WBS, RowType.UNKNOWN}:
            forced_wbs = row_type == RowType.WBS and decision.row_type != RowType.WBS
            level_override = cfg.wbs_level_overrides.get(extracted.source_row)
            if builder.accept_heading(extracted, forced=forced_wbs, level_override=level_override):
                result.wbs_rows += 1
            result.skipped_rows += 1
        elif extracted.amount is None or extracted.amount <= 0:
            result.skipped_rows += 1
        else:
            item_counter += 1
            output_ws.append(
                builder.build_item(
                    extracted,
                    cfg.sheet_name,
                    wbs_no=f"{project_number}.{item_counter}",
                )
            )
            result.item_rows += 1
            result.output_rows += 1

        if progress:
            progress(index, total_rows, f"{cfg.sheet_name}: row {extracted.source_row}")

    style_output_sheet(output_ws)
    if result.output_rows == 0:
        result.status = "WARNING"
        result.message = "No rows with Amount greater than zero were produced. Check the mapping."
    return result


def normalize_sheet_from_tree(
    source_ws,
    output_ws,
    cfg: SheetConfig,
    draft_tree: DraftTree,
    progress: ProgressCallback | None = None,
    *,
    project_number: int = 1,
) -> SheetResult:
    """Export one sheet using the approved Draft Tree as source of truth.

    ``project_number`` is the sheet's stable position in the selected-project
    order. Therefore each sheet owns one unique number branch: 1.x, 2.x, 3.x,
    and so on. WBS numbers are regenerated from tree order at export time;
    internal Node IDs remain stable in the session even when the user moves or
    indents rows.
    """

    result = SheetResult(sheet_name=cfg.sheet_name, status="OK")
    extractor = SheetExtractor()
    extracted = {row.source_row: row for row in extractor.extract(source_ws, cfg)}
    result.source_rows = len(extracted)

    output_ws.append(OUTPUT_HEADERS)
    _style_header(output_ws)

    # Export the sheet itself as the WBS-2 root. This makes the first number of
    # each normalized sheet visible and keeps the merged Project sheet unique.
    output_ws.append([
        str(project_number),
        "Project",
        cfg.sheet_name,
        "",
        "",
        cfg.sheet_name,
        "",
        "",
        "",
        "",
        "",
    ])
    result.wbs_rows += 1
    result.output_rows += 1

    total_top_nodes = max(1, len(draft_tree.root.children))

    def visit(
        node: DraftTreeNode,
        number: str,
        *,
        wbs3: str = "",
        wbs4: str = "",
        wbs_depth: int = 0,
    ) -> None:
        next_wbs3, next_wbs4 = wbs3, wbs4
        next_depth = wbs_depth

        if node.is_wbs:
            next_depth += 1
            if next_depth == 1:
                next_wbs3 = node.description
                next_wbs4 = ""
            else:
                next_wbs4 = node.description

            output_ws.append([
                number,
                "Project",
                cfg.sheet_name,
                next_wbs3,
                next_wbs4,
                node.description,
                "",
                "",
                "",
                "",
                "",
            ])
            result.wbs_rows += 1
            result.output_rows += 1

        elif node.node_type == TreeNodeType.ITEM:
            row = extracted.get(node.source_row or -1)
            if row is not None and row.amount is not None and row.amount > 0:
                output_ws.append([
                    number,
                    "Project",
                    cfg.sheet_name,
                    wbs3,
                    wbs4,
                    node.description or row.description,
                    row.unit,
                    row.qty,
                    row.material,
                    row.labor,
                    row.amount,
                ])
                result.item_rows += 1
                result.output_rows += 1
            else:
                result.skipped_rows += 1
        else:
            result.skipped_rows += 1

        for child_index, child in enumerate(node.children, start=1):
            visit(
                child,
                f"{number}.{child_index}",
                wbs3=next_wbs3,
                wbs4=next_wbs4,
                wbs_depth=next_depth,
            )

    for index, child in enumerate(draft_tree.root.children, start=1):
        visit(child, f"{project_number}.{index}")
        if progress:
            progress(index, total_top_nodes, f"{cfg.sheet_name}: tree node {index}")

    style_output_sheet(output_ws)
    if result.item_rows == 0:
        result.status = "WARNING"
        result.message = "No tree ITEM rows with Amount greater than zero were produced."
    return result


def normalized_preview_rows(source_ws, cfg: SheetConfig, limit: int = 60) -> list[list[object]]:
    rows: list[list[object]] = []
    extractor = SheetExtractor()
    classifier = RowClassifier()
    builder = WBSBuilder(cfg.sheet_name)
    extracted_rows = list(extractor.extract(source_ws, cfg))
    decisions = classifier.classify_rows(extracted_rows)
    item_counter = 0
    for extracted, decision in zip(extracted_rows, decisions, strict=True):
        row_type = resolve_row_type(decision, extracted.source_row, cfg.classification_overrides)
        if row_type in {RowType.EMPTY, RowType.TOTAL, RowType.NOTE, RowType.DETAIL, RowType.IGNORE}:
            continue
        if row_type in {RowType.WBS, RowType.UNKNOWN}:
            forced_wbs = row_type == RowType.WBS and decision.row_type != RowType.WBS
            builder.accept_heading(
                extracted,
                forced=forced_wbs,
                level_override=cfg.wbs_level_overrides.get(extracted.source_row),
            )
            continue
        if extracted.amount is None or extracted.amount <= 0:
            continue
        item_counter += 1
        rows.append(
            builder.build_item(
                extracted,
                cfg.sheet_name,
                wbs_no=f"1.{item_counter}",
            )
        )
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
            draft_tree = request.draft_trees.get(cfg.sheet_name)
            if draft_tree is not None:
                result = normalize_sheet_from_tree(
                    source_wb[cfg.sheet_name],
                    target_ws,
                    cfg,
                    draft_tree,
                    sheet_progress,
                    project_number=sheet_index,
                )
            else:
                result = normalize_sheet(
                    source_wb[cfg.sheet_name],
                    target_ws,
                    cfg,
                    sheet_progress,
                    project_number=sheet_index,
                )
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


def _wbs_depth(value: object) -> int:
    text = str(value or "").strip()
    if not text:
        return 0
    return len([part for part in text.split(".") if part])


def _apply_outline_and_row_styles(ws) -> None:
    """Create Excel Outline levels directly from the exported WBS numbering."""
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.applyStyles = True
    for row in range(2, ws.max_row + 1):
        depth = _wbs_depth(ws.cell(row, 1).value)
        # The sheet root stays visible at level 0. Descendants are grouped under it.
        ws.row_dimensions[row].outlineLevel = min(max(depth - 1, 0), 7)
        amount = ws.cell(row, 11).value
        is_wbs = amount in (None, "")
        if is_wbs:
            fill = "D9EAF7" if depth <= 2 else "EAF2F8"
            for cell in ws[row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor=fill)


def style_output_sheet(ws, *, apply_outline: bool = True) -> None:
    widths = {
        "A": 16,
        "B": 22,
        "C": 24,
        "D": 30,
        "E": 30,
        "F": 52,
        "G": 12,
        "H": 14,
        "I": 18,
        "J": 18,
        "K": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    ws.print_area = f"A1:K{max(1, ws.max_row)}"
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.2, footer=0.2)
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in (8, 9, 10, 11):
        for cells in ws.iter_cols(min_col=col, max_col=col, min_row=2):
            cells[0].number_format = '#,##0.00;[Red]-#,##0.00;'
    if apply_outline:
        _apply_outline_and_row_styles(ws)


def excel_sheet_ref(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def add_project_sheet(wb, normalized_sheet_names: list[str]) -> int:
    """Create the formula-linked merged sheet.

    WBS No. values are already project-wide and unique in the source sheets, so
    the Project sheet only needs transparent Excel links; no renumbering or
    lookup keys are required during merge.
    """

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
            ws.row_dimensions[project_row].outlineLevel = source_ws.row_dimensions[source_row].outlineLevel
            ws.row_dimensions[project_row].hidden = source_ws.row_dimensions[source_row].hidden
            # Project formulas have not been calculated yet, so copy the visible
            # WBS row style from the linked normalized sheet.
            if source_ws.cell(source_row, 11).value in (None, ""):
                for column in range(1, len(OUTPUT_HEADERS) + 1):
                    ws.cell(project_row, column)._style = source_ws.cell(source_row, column)._style
            project_rows += 1
    _style_header(ws)
    style_output_sheet(ws, apply_outline=False)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.applyStyles = True
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
        ws.cell(row_no, 2, f'=COUNT({ref}!K2:K{last_row})')
        ws.cell(row_no, 3, f'=SUM({ref}!K2:K{last_row})')
    last_detail_row = ws.max_row
    grand_row = last_detail_row + 2
    ws.cell(grand_row, 1, "Grand Total")
    ws.cell(grand_row, 2, f"=SUM(B{first_detail_row}:B{last_detail_row})" if last_detail_row >= first_detail_row else 0)
    ws.cell(grand_row, 3, f"=SUM(C{first_detail_row}:C{last_detail_row})" if last_detail_row >= first_detail_row else 0)
    project_row = grand_row + 2
    project_last_row = max(2, wb["Project"].max_row)
    ws.cell(project_row, 1, "Project Sheet Total")
    ws.cell(project_row, 3, f"=SUM('Project'!K2:K{project_last_row})")
    ws.cell(project_row + 1, 1, "Difference")
    ws.cell(project_row + 1, 3, f"=C{project_row}-C{grand_row}")
    for cell in ws[1]:
        cell.font = Font(size=14, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[grand_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="548235")
    ws.freeze_panes = "A3"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 20
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
        ws.append([
            result.sheet_name,
            result.status,
            result.source_rows,
            result.output_rows,
            result.wbs_rows,
            result.item_rows,
            result.skipped_rows,
            result.message,
        ])
    for cell in ws[1]:
        cell.font = Font(size=14, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[6]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A7"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    for col in range(3, 8):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["H"].width = 50
