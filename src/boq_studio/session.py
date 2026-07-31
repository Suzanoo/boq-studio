from __future__ import annotations

import hashlib
import json
import os
import sys
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .core.tree import DraftTree, DraftTreeNode, TreeNodeType
from .models import SheetConfig

SESSION_VERSION = 1


def workbook_fingerprint(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def session_directory() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
        return base / "BOQ Studio" / "sessions"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "BOQ Studio" / "sessions"
    base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    return base / "boq-studio" / "sessions"


def session_path_for(workbook: Path, directory: Path | None = None) -> Path:
    key = hashlib.sha256(str(workbook.resolve()).encode("utf-8")).hexdigest()
    return (directory or session_directory()) / f"{key}.json"


def node_to_dict(node: DraftTreeNode) -> dict[str, Any]:
    return {
        "node_id": node.node_id,
        "node_type": node.node_type.value,
        "description": node.description,
        "source_row": node.source_row,
        "item_no": node.item_no,
        "confidence": node.confidence,
        "reason": node.reason,
        "amount": node.amount,
        "children": [node_to_dict(child) for child in node.children],
    }


def node_from_dict(data: dict[str, Any]) -> DraftTreeNode:
    return DraftTreeNode(
        node_id=data["node_id"],
        node_type=TreeNodeType(data["node_type"]),
        description=data.get("description", ""),
        source_row=data.get("source_row"),
        item_no=data.get("item_no", ""),
        confidence=float(data.get("confidence", 1.0)),
        reason=data.get("reason", ""),
        amount=data.get("amount"),
        children=[node_from_dict(child) for child in data.get("children", [])],
    )


def tree_to_dict(tree: DraftTree) -> dict[str, Any]:
    return {"sheet_name": tree.sheet_name, "root": node_to_dict(tree.root)}


def tree_from_dict(data: dict[str, Any]) -> DraftTree:
    return DraftTree(data["sheet_name"], node_from_dict(data["root"]))


def config_to_dict(config: SheetConfig) -> dict[str, Any]:
    return asdict(config)


def config_from_dict(data: dict[str, Any]) -> SheetConfig:
    values = dict(data)
    values["classification_overrides"] = {
        int(key): value for key, value in values.get("classification_overrides", {}).items()
    }
    values["wbs_level_overrides"] = {
        int(key): int(value) for key, value in values.get("wbs_level_overrides", {}).items()
    }
    return SheetConfig(**values)


def build_session_payload(
    workbook: Path,
    selected: dict[str, bool],
    configs: dict[str, SheetConfig],
    trees: dict[str, DraftTree],
    tree_ready: set[str],
    exported: set[str],
) -> dict[str, Any]:
    stat = workbook.stat()
    return {
        "version": SESSION_VERSION,
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "source": {
            "path": str(workbook.resolve()),
            "sha256": workbook_fingerprint(workbook),
            "size": stat.st_size,
            "modified_ns": stat.st_mtime_ns,
        },
        "selected": selected,
        "configs": {name: config_to_dict(config) for name, config in configs.items()},
        "trees": {name: tree_to_dict(tree) for name, tree in trees.items()},
        "tree_ready": sorted(tree_ready),
        "exported": sorted(exported),
    }


def save_session(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def load_session(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("version") != SESSION_VERSION:
        raise ValueError("Unsupported BOQ Studio session version.")
    return payload


def session_matches_workbook(payload: dict[str, Any], workbook: Path) -> bool:
    return payload.get("source", {}).get("sha256") == workbook_fingerprint(workbook)
