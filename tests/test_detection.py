from pathlib import Path

from openpyxl import Workbook

from boq_studio.detection import auto_detect_config
from boq_studio.models import SheetConfig
from boq_studio.normalize import normalized_preview_rows
from boq_studio.templates import load_template, save_template


def test_auto_detect_mapping_and_preview(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Arch"
    ws.append(["Project BOQ"])
    ws.append([])
    ws.append(["ลำดับ", "รายละเอียด", "หน่วย", "ปริมาณ", "ค่าวัสดุ", "ค่าแรง", "จำนวนเงิน"])
    ws.append(["1", "งานสถาปัตยกรรม", None, None, None, None, None])
    ws.append(["1.1", "งานพื้น", None, None, None, None, None])
    ws.append(["1", "กระเบื้องพื้น", "ตร.ม.", 100, 20000, 5000, 25000])

    cfg = auto_detect_config(ws, "Arch")
    assert cfg.header_row == 3
    assert cfg.start_row == 4
    assert cfg.item_no_col == 1
    assert cfg.description_cols == [2]
    assert cfg.unit_col == 3
    assert cfg.qty_col == 4
    assert cfg.material_col == 5
    assert cfg.labor_col == 6
    assert cfg.amount_col == 7

    rows = normalized_preview_rows(ws, cfg)
    assert len(rows) == 1
    assert rows[0][2] == "Arch"
    assert rows[0][5] == "กระเบื้องพื้น"
    assert rows[-1][10] == 25000


def test_save_and_load_template(tmp_path: Path):
    path = tmp_path / "mapping.json"
    cfg = SheetConfig(sheet_name="Old", header_row=6, start_row=33, description_cols=[3, 4], qty_col=5, amount_col=13)
    save_template(path, cfg)
    loaded = load_template(path, sheet_name="New")
    assert loaded.sheet_name == "New"
    assert loaded.header_row == 6
    assert loaded.description_cols == [3, 4]
    assert loaded.amount_col == 13
