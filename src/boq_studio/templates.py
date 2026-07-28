from __future__ import annotations

import json
import re
from dataclasses import asdict
from pathlib import Path

from .models import SheetConfig


def template_directory() -> Path:
    path = Path.home() / ".boq_studio" / "mappings"
    path.mkdir(parents=True, exist_ok=True)
    return path


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._") or "mapping"


def default_template_path(workbook_path: Path, sheet_name: str) -> Path:
    return template_directory() / f"{safe_filename(workbook_path.stem)}__{safe_filename(sheet_name)}.json"


def save_template(path: Path, cfg: SheetConfig) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"version": 1, "mapping": asdict(cfg)}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_template(path: Path, sheet_name: str | None = None) -> SheetConfig:
    payload = json.loads(path.read_text(encoding="utf-8"))
    data = payload.get("mapping", payload)
    if sheet_name:
        data["sheet_name"] = sheet_name
    return SheetConfig(**data)
