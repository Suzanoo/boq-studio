from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from collections.abc import Iterable, Sequence

from .classification import (
    ClassificationResult,
    RowClassifier,
    RowType,
    resolve_row_type,
)
from .extraction import ExtractedRow
from .item_numbering import ItemNumber, ItemNumberParser


class TreeNodeType(StrEnum):
    SHEET_ROOT = "SHEET_ROOT"
    WBS = "WBS"
    ITEM = "ITEM"
    DETAIL = "DETAIL"
    TOTAL = "TOTAL"
    UNKNOWN = "UNKNOWN"
    VIRTUAL_WBS = "VIRTUAL_WBS"


_ROW_TO_NODE = {
    RowType.WBS: TreeNodeType.WBS,
    RowType.ITEM: TreeNodeType.ITEM,
    RowType.DETAIL: TreeNodeType.DETAIL,
    RowType.TOTAL: TreeNodeType.TOTAL,
    RowType.UNKNOWN: TreeNodeType.UNKNOWN,
}


@dataclass(slots=True)
class DraftTreeNode:
    node_id: str
    node_type: TreeNodeType
    description: str
    source_row: int | None = None
    item_no: str = ""
    confidence: float = 1.0
    reason: str = ""
    amount: float | None = None
    children: list["DraftTreeNode"] = field(default_factory=list)

    def add_child(self, node: "DraftTreeNode") -> None:
        self.children.append(node)

    @property
    def is_wbs(self) -> bool:
        return self.node_type in {TreeNodeType.WBS, TreeNodeType.VIRTUAL_WBS}


@dataclass(slots=True)
class DraftTree:
    sheet_name: str
    root: DraftTreeNode

    def walk(self) -> Iterable[tuple[int, DraftTreeNode]]:
        """Yield ``(depth, node)`` in display order, excluding the sheet root."""

        def visit(parent: DraftTreeNode, depth: int):
            for child in parent.children:
                yield depth, child
                yield from visit(child, depth + 1)

        yield from visit(self.root, 1)

    def find(self, node_id: str) -> DraftTreeNode | None:
        if self.root.node_id == node_id:
            return self.root
        for _, node in self.walk():
            if node.node_id == node_id:
                return node
        return None


@dataclass(slots=True)
class _WBSContext:
    node: DraftTreeNode
    number: ItemNumber | None
    indent_level: int


class DraftTreeBuilder:
    """Build a first-pass WBS tree from classified rows.

    The builder uses Item No. ancestry when available and falls back to source
    order plus Excel indentation. It deliberately produces a *draft*: ambiguous
    rows remain visible as UNKNOWN nodes so the later tree editor can fix them
    without losing source information.
    """

    def __init__(
        self,
        classifier: RowClassifier | None = None,
        item_number_parser: ItemNumberParser | None = None,
    ) -> None:
        self.item_number_parser = item_number_parser or ItemNumberParser()
        self.classifier = classifier or RowClassifier(self.item_number_parser)

    def build(
        self,
        sheet_name: str,
        rows: Sequence[ExtractedRow],
        decisions: Sequence[ClassificationResult] | None = None,
        overrides: dict[int, str] | None = None,
    ) -> DraftTree:
        if decisions is None:
            decisions = self.classifier.classify_rows(rows)
        if len(rows) != len(decisions):
            raise ValueError("rows and decisions must have the same length")

        root = DraftTreeNode(
            node_id=f"sheet:{sheet_name}",
            node_type=TreeNodeType.SHEET_ROOT,
            description=sheet_name,
        )
        tree = DraftTree(sheet_name=sheet_name, root=root)
        contexts: list[_WBSContext] = []
        last_item: DraftTreeNode | None = None

        for row, decision in zip(rows, decisions, strict=True):
            row_type = resolve_row_type(decision, row.source_row, overrides)
            if row_type in {RowType.EMPTY, RowType.NOTE, RowType.IGNORE}:
                continue

            node_type = _ROW_TO_NODE.get(row_type, TreeNodeType.UNKNOWN)
            node = DraftTreeNode(
                node_id=f"{sheet_name}:{row.source_row}",
                node_type=node_type,
                description=row.description,
                source_row=row.source_row,
                item_no=row.item_no,
                confidence=decision.confidence,
                reason=decision.reason,
                amount=row.amount,
            )

            if row_type == RowType.WBS:
                number = self.item_number_parser.parse(row.item_no)
                parent, retained = self._wbs_parent(root, contexts, number, row.indent_level)
                parent.add_child(node)
                contexts = [*retained, _WBSContext(node, number, row.indent_level)]
                last_item = None
                continue

            if row_type == RowType.ITEM:
                parent = self._item_parent(root, contexts, row)
                parent.add_child(node)
                last_item = node
                continue

            if row_type == RowType.DETAIL and last_item is not None:
                last_item.add_child(node)
                continue

            self._current_parent(root, contexts).add_child(node)
            if row_type not in {RowType.DETAIL, RowType.TOTAL}:
                last_item = None

        return tree

    def _wbs_parent(
        self,
        root: DraftTreeNode,
        contexts: list[_WBSContext],
        number: ItemNumber | None,
        indent_level: int,
    ) -> tuple[DraftTreeNode, list[_WBSContext]]:
        if number is not None:
            for index in range(len(contexts) - 1, -1, -1):
                context = contexts[index]
                if context.number is not None and context.number.is_parent_of(number):
                    return context.node, contexts[: index + 1]
            return root, []

        # Text-only WBS headings use visual indentation as a conservative
        # fallback. Equal indentation creates a sibling, greater indentation a
        # child. With no useful indentation, headings remain sheet-level WBS.
        for index in range(len(contexts) - 1, -1, -1):
            context = contexts[index]
            if context.indent_level < indent_level:
                return context.node, contexts[: index + 1]
        return root, []

    def _item_parent(
        self,
        root: DraftTreeNode,
        contexts: list[_WBSContext],
        row: ExtractedRow,
    ) -> DraftTreeNode:
        item_number = self.item_number_parser.parse(row.item_no)
        if item_number is not None:
            for context in reversed(contexts):
                if context.number is not None and context.number.is_parent_of(item_number):
                    return context.node
        return self._current_parent(root, contexts)

    @staticmethod
    def _current_parent(root: DraftTreeNode, contexts: list[_WBSContext]) -> DraftTreeNode:
        return contexts[-1].node if contexts else root
