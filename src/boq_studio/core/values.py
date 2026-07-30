from __future__ import annotations

import math
import re
from typing import Any


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return " ".join(value.replace("\n", " ").split())
    return str(value).strip()


def safe_number(value: Any) -> float | None:
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = safe_text(value).replace(",", "").replace("฿", "").replace("บาท", "").strip()
    if text in {"-", "–", "—"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def is_blank_value(value: Any) -> bool:
    return value is None or safe_text(value) == ""


def numbering_level(text: str) -> int | None:
    text = text.strip()
    match = re.fullmatch(r"(\d+(?:\.\d+){0,3})", text)
    if not match:
        match = re.match(r"^(\d+(?:\.\d+){0,3})[)\-\s]", text)
    if match:
        return min(4, match.group(1).count(".") + 1)
    if re.match(r"^(?:หมวด|ส่วนที่|งานที่)\s*(\d+)", text, flags=re.I):
        return 1
    return None
