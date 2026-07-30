from __future__ import annotations

from dataclasses import dataclass
from uuid import uuid4

from .tree import DraftTree, DraftTreeNode, TreeNodeType


class TreeEditError(ValueError):
    """Raised when a requested tree edit is not valid."""


@dataclass(slots=True)
class TreeLocation:
    parent: DraftTreeNode
    index: int


class DraftTreeEditor:
    """Small, UI-independent edit service for a :class:`DraftTree`.

    MS-5.4 intentionally keeps the rules conservative. Only WBS nodes can own
    siblings through indentation, the sheet root cannot be moved, and source
    nodes cannot be deleted. Virtual WBS nodes may be removed without losing
    their children.
    """

    def __init__(self, tree: DraftTree, max_wbs_depth: int = 2) -> None:
        self.tree = tree
        self.max_wbs_depth = max_wbs_depth

    def locate(self, node_id: str) -> TreeLocation | None:
        def visit(parent: DraftTreeNode) -> TreeLocation | None:
            for index, child in enumerate(parent.children):
                if child.node_id == node_id:
                    return TreeLocation(parent, index)
                found = visit(child)
                if found is not None:
                    return found
            return None

        return visit(self.tree.root)

    def depth(self, node_id: str) -> int:
        for depth, node in self.tree.walk():
            if node.node_id == node_id:
                return depth
        raise TreeEditError("Tree node was not found.")

    def indent(self, node_id: str) -> None:
        location = self._required_location(node_id)
        if location.index == 0:
            raise TreeEditError("The first row has no previous sibling to indent under.")
        previous = location.parent.children[location.index - 1]
        if not previous.is_wbs:
            raise TreeEditError("Indent requires the previous sibling to be a WBS node.")
        if self.depth(previous.node_id) >= self.max_wbs_depth and self.tree.find(node_id).is_wbs:
            raise TreeEditError("Maximum supported WBS depth is WBS-4.")
        node = location.parent.children.pop(location.index)
        previous.children.append(node)

    def outdent(self, node_id: str) -> None:
        location = self._required_location(node_id)
        parent_location = self.locate(location.parent.node_id)
        if parent_location is None:
            raise TreeEditError("The row is already at the sheet root.")
        node = location.parent.children.pop(location.index)
        parent_location.parent.children.insert(parent_location.index + 1, node)

    def create_wbs(self, description: str, selected_id: str | None = None) -> DraftTreeNode:
        description = description.strip()
        if not description:
            raise TreeEditError("WBS description cannot be blank.")
        node = DraftTreeNode(
            node_id=f"virtual:{self.tree.sheet_name}:{uuid4().hex}",
            node_type=TreeNodeType.VIRTUAL_WBS,
            description=description,
            confidence=1.0,
            reason="Created by user",
        )
        if not selected_id:
            self.tree.root.children.append(node)
            return node

        selected = self.tree.find(selected_id)
        if selected is None:
            raise TreeEditError("Selected tree node was not found.")
        if selected.is_wbs:
            if self.depth(selected.node_id) >= self.max_wbs_depth:
                raise TreeEditError("Maximum supported WBS depth is WBS-4.")
            selected.children.insert(0, node)
            return node

        location = self._required_location(selected_id)
        location.parent.children.insert(location.index, node)
        return node

    def rename_wbs(self, node_id: str, description: str) -> None:
        node = self.tree.find(node_id)
        if node is None or not node.is_wbs:
            raise TreeEditError("Only WBS nodes can be renamed.")
        description = description.strip()
        if not description:
            raise TreeEditError("WBS description cannot be blank.")
        node.description = description

    def delete_virtual_wbs(self, node_id: str) -> None:
        location = self._required_location(node_id)
        node = location.parent.children[location.index]
        if node.node_type != TreeNodeType.VIRTUAL_WBS:
            raise TreeEditError("Only a user-created WBS can be deleted.")
        location.parent.children[location.index: location.index + 1] = node.children

    def _required_location(self, node_id: str) -> TreeLocation:
        location = self.locate(node_id)
        if location is None:
            raise TreeEditError("Tree node was not found.")
        return location
