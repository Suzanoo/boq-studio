from __future__ import annotations

import hashlib
import json
import math
import os
import sys
import tempfile
from contextlib import contextmanager
from uuid import uuid4
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .core.classification import RowType
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


def save_session(path: Path, payload: dict[str, Any], *, expected_token=...) -> str:
    """Replace safely, retaining the last valid disk state; legacy callers still work."""
    validate_session(payload)
    data = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8")
    with _session_lock(path):
        if expected_token is not ... and session_token(path) != expected_token:
            raise SessionConflictError("Saved session changed in another window; reopen before saving.")
        if path.exists():
            old = path.read_bytes()
            try:
                previous = json.loads(old)
                validate_session(previous)
            except Exception:
                # Do not replace a useful .bak with malformed primary bytes.
                _archive_bytes(path, old)
            else:
                comparable = lambda value: {k: v for k, v in value.items() if k != "saved_at"}
                if comparable(previous) == comparable(json.loads(data)):
                    return hashlib.sha256(old).hexdigest()
                _atomic_write(path.with_suffix(path.suffix + ".bak"), old)
        _atomic_write(path, data)
    return hashlib.sha256(data).hexdigest()


def load_session(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    validate_session(payload)
    return payload


def session_matches_workbook(payload: dict[str, Any], workbook: Path) -> bool:
    return payload.get("source", {}).get("sha256") == workbook_fingerprint(workbook)


class SessionConflictError(ValueError):
    """The source or saved session changed since this window opened it."""


def session_token(path: Path) -> str | None:
    """A missing file is a meaningful expected version, not 'ignore conflicts'."""
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except FileNotFoundError:
        return None


def validate_session(payload: dict[str, Any]) -> None:
    """Validate before activation or replacement; keep the version-1 format."""
    if not isinstance(payload, dict) or payload.get("version") != SESSION_VERSION:
        raise ValueError("Unsupported BOQ Studio session version.")
    source = payload.get("source")
    if not isinstance(source, dict) or not isinstance(source.get("path"), str):
        raise ValueError("Session source is missing.")
    digest = source.get("sha256", "")
    if not isinstance(digest, str) or len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
        raise ValueError("Invalid session source fingerprint.")
    for key in ("selected", "configs", "trees"):
        if not isinstance(payload.get(key), dict):
            raise ValueError(f"Invalid session {key}.")
        if any(not isinstance(name, str) or not name for name in payload[key]):
            raise ValueError(f"Invalid sheet name in {key}.")
    if any(type(value) is not bool for value in payload["selected"].values()):
        raise ValueError("Invalid sheet selection.")
    for name, raw in payload["configs"].items():
        config = config_from_dict(raw)
        if config.sheet_name != name:
            raise ValueError("Mapping sheet name does not match session key.")
        for value in (config.header_row, config.start_row, config.max_wbs_level):
            if type(value) is not int or value < 1:
                raise ValueError("Invalid mapping row/level.")
        if not isinstance(config.description_cols, list):
            raise ValueError("Invalid description columns.")
        cols = [*config.description_cols, config.item_no_col, config.unit_col,
                config.qty_col, config.material_col, config.labor_col, config.amount_col]
        if any(value is not None and (type(value) is not int or value < 1) for value in cols):
            raise ValueError("Invalid mapping column.")
        if any(row < 1 or value not in {v.value for v in RowType}
               for row, value in config.classification_overrides.items()):
            raise ValueError("Invalid classification override.")
        if any(row < 1 or value not in {2, 3, 4}
               for row, value in config.wbs_level_overrides.items()):
            raise ValueError("Invalid WBS level override.")
    for name, raw in payload["trees"].items():
        if name not in payload["configs"] or raw.get("sheet_name") != name:
            raise ValueError("Tree has no matching mapping/sheet.")
        root = raw["root"]
        if root.get("node_type") != TreeNodeType.SHEET_ROOT.value:
            raise ValueError("Invalid tree root.")
        ids: set[str] = set()
        stack = [(root, 0)]
        while stack:
            node, depth = stack.pop()
            if depth > 200:
                raise ValueError("Tree nesting exceeds the supported session limit.")
            node_id = node.get("node_id")
            if not isinstance(node_id, str) or not node_id or node_id in ids:
                raise ValueError("Missing or duplicate tree node ID.")
            ids.add(node_id)
            kind = TreeNodeType(node["node_type"])
            if depth and kind == TreeNodeType.SHEET_ROOT:
                raise ValueError("Nested sheet root.")
            if not isinstance(node.get("description", ""), str):
                raise ValueError("Invalid node description.")
            row = node.get("source_row")
            if row is not None and (type(row) is not int or row < 1):
                raise ValueError("Invalid source row.")
            for key in ("amount", "confidence"):
                value = node.get(key)
                if value is not None and (type(value) not in (int, float) or not math.isfinite(value)):
                    raise ValueError(f"Invalid node {key}.")
            children = node.get("children", [])
            if not isinstance(children, list):
                raise ValueError("Invalid node children.")
            stack.extend((child, depth + 1) for child in children)
    for key in ("tree_ready", "exported"):
        if not isinstance(payload.get(key), list) or any(not isinstance(n, str) for n in payload[key]):
            raise ValueError(f"Invalid {key} state.")
        if not set(payload[key]) <= payload["trees"].keys():
            raise ValueError(f"{key} refers to a missing tree.")
    if any(not payload["configs"][name].get("description_cols") for name in payload["tree_ready"]):
        raise ValueError("Approved tree has no Description mapping.")
    if not set(payload["exported"]) <= set(payload["tree_ready"]):
        raise ValueError("Exported state refers to an unapproved tree.")


# The lock is a stable sibling, not the file being atomically replaced.
# It carries no project data and OS locks are released when a process exits.
@contextmanager
def _session_lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.with_suffix(path.suffix + ".lock").open("a+b") as handle:
        if os.name == "nt":
            import msvcrt
            if handle.tell() == 0:
                handle.write(b"0")
                handle.flush()
            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise SessionConflictError("Another BOQ Studio window is saving this session.") from exc
            try:
                yield
            finally:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as exc:
                raise SessionConflictError("Another BOQ Studio window is saving this session.") from exc
            try:
                yield
            finally:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _atomic_write(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    temporary = Path(temporary)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        if temporary.read_bytes() != data:
            raise OSError("Session write verification failed.")
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def _archive_bytes(path: Path, data: bytes) -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    target = path.with_name(f"{path.stem}.recovery-{stamp}-{uuid4().hex}.json")
    _atomic_write(target, data)
    return target


_UNCHECKED = object()


def archive_session(path: Path, *, expected_token: str | None | object = _UNCHECKED) -> Path:
    """Retain an immutable copy before an explicitly destructive lifecycle action."""
    with _session_lock(path):
        if expected_token is not _UNCHECKED and session_token(path) != expected_token:
            raise SessionConflictError("Saved session changed in another window; reopen before continuing.")
        return _archive_bytes(path, path.read_bytes())


