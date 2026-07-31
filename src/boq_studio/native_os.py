from __future__ import annotations

import json
import os
from pathlib import Path

APP_NAME = "BOQ Studio"
MAX_RECENT_FILES = 8


def is_macos(system_name: str) -> bool:
    return system_name.casefold() == "darwin"


def user_config_dir(system_name: str, home: Path | None = None) -> Path:
    home = home or Path.home()
    name = system_name.casefold()
    if name == "windows":
        base = Path(os.environ.get("APPDATA", home / "AppData" / "Roaming"))
        return base / APP_NAME
    if name == "darwin":
        return home / "Library" / "Application Support" / APP_NAME
    return Path(os.environ.get("XDG_CONFIG_HOME", home / ".config")) / "boq-studio"


def recent_files_path(system_name: str, home: Path | None = None) -> Path:
    return user_config_dir(system_name, home) / "recent_files.json"


def layout_state_path(system_name: str, home: Path | None = None) -> Path:
    return user_config_dir(system_name, home) / "layout.json"


def view_state_path(system_name: str, home: Path | None = None) -> Path:
    return user_config_dir(system_name, home) / "view_state.json"


def normalize_recent_files(paths: list[str | Path], limit: int = MAX_RECENT_FILES) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for raw in paths:
        path = Path(raw).expanduser()
        key = str(path.resolve(strict=False)).casefold()
        if key in seen or not path.exists() or path.suffix.casefold() != ".xlsx":
            continue
        seen.add(key)
        result.append(path)
        if len(result) >= limit:
            break
    return result


def load_recent_files(path: Path) -> list[Path]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    return normalize_recent_files(data)


def save_recent_files(path: Path, files: list[str | Path]) -> list[Path]:
    normalized = normalize_recent_files(files)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps([str(item) for item in normalized], indent=2), encoding="utf-8")
    return normalized


def add_recent_file(path: Path, files: list[str | Path]) -> list[Path]:
    return normalize_recent_files([path, *files])
