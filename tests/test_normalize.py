from pathlib import Path

from openpyxl import Workbook, load_workbook

from boq_studio.models import NormalizeRequest, SheetConfig
from boq_studio.normalize import normalize_workbook


def test_normalize_workbook(tmp_path: Path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "source_OKD.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Breakdown"
    ws.append(["No", "Description", "Unit", "Qty", "Material", "Labor", "Amount"])
    ws.append([1, "งานโครงสร้าง", None, None, None, None, None])
    ws.append(["1.1", "งานคอนกรีต", None, None, None, None, None])
    ws.append([1, "คอนกรีต 240 ksc", "m3", 10, 20000, 5000, 25000])
    ws.append([None, "รวม", None, None, None, None, 25000])
    wb.save(source)

    request = NormalizeRequest(
        input_path=source,
        output_path=output,
        sheets=[SheetConfig(
            sheet_name="Breakdown",
            header_row=1,
            start_row=2,
            item_no_col=1,
            description_cols=[2],
            unit_col=3,
            qty_col=4,
            material_col=5,
            labor_col=6,
            amount_col=7,
        )],
    )
    results = normalize_workbook(request)
    assert results[0].output_rows == 3
    assert results[0].wbs_rows == 2
    assert results[0].item_rows == 1

    out = load_workbook(output, data_only=True)
    assert "LOG" in out.sheetnames
    assert "Breakdown" in out.sheetnames
    ws_out = out["Breakdown"]
    assert ws_out["A2"].value == "1"
    assert ws_out["A3"].value == "1.1"
    assert ws_out["C4"].value == "ITEM"
    assert ws_out["M4"].value == 25000
