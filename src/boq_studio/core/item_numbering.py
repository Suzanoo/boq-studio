from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
import re


class NumberingStyle(StrEnum):
    NUMERIC = "NUMERIC"
    ALPHA = "ALPHA"


@dataclass(slots=True, frozen=True)
class ItemNumber:
    """A normalized hierarchical Item No. such as ``1.2`` or ``A.1``.

    Product/model codes are deliberately rejected. For example ``F1.1`` and
    ``FDX-301`` are not hierarchy numbers because the first segment combines
    letters and digits or uses a hyphenated code shape.
    """

    raw: str
    style: NumberingStyle
    path: tuple[str, ...]

    @property
    def level(self) -> int:
        return len(self.path)

    @property
    def parent_path(self) -> tuple[str, ...] | None:
        return self.path[:-1] or None

    def is_parent_of(self, other: "ItemNumber") -> bool:
        return (
            self.style == other.style
            and len(other.path) > len(self.path)
            and other.path[: len(self.path)] == self.path
        )

    def is_sibling_of(self, other: "ItemNumber") -> bool:
        return (
            self.style == other.style
            and self.level == other.level
            and self.parent_path == other.parent_path
        )


class ItemNumberParser:
    """Parse common BOQ hierarchy numbering without guessing product codes."""

    _WRAPPERS = re.compile(r"^[\s\(\[]+|[\s\)\].:]+$")
    _NUMERIC = re.compile(r"^\d+(?:\.\d+){0,5}$")
    _ALPHA = re.compile(r"^[A-Za-z](?:\.\d+){0,5}$")

    def parse(self, value: object) -> ItemNumber | None:
        if value is None:
            return None
        raw = str(value).strip()
        if not raw:
            return None

        normalized = self._strip_wrappers(raw)
        if not normalized or "-" in normalized or "/" in normalized:
            return None

        if self._NUMERIC.fullmatch(normalized):
            return ItemNumber(
                raw=raw,
                style=NumberingStyle.NUMERIC,
                path=tuple(str(int(part)) for part in normalized.split(".")),
            )

        if self._ALPHA.fullmatch(normalized):
            parts = normalized.split(".")
            return ItemNumber(
                raw=raw,
                style=NumberingStyle.ALPHA,
                path=(parts[0].upper(), *(str(int(part)) for part in parts[1:])),
            )

        return None

    @classmethod
    def _strip_wrappers(cls, value: str) -> str:
        # Common BOQ forms: 1), (1), 1.1), A., A.1:
        cleaned = value.strip()
        cleaned = re.sub(r"^[\(\[]", "", cleaned)
        cleaned = re.sub(r"[\)\]]$", "", cleaned)
        cleaned = re.sub(r"[:.]$", "", cleaned)
        return cleaned.strip()
