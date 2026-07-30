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

    The sheet root cannot be edited. Any other node may be reordered or
    removed. Deleting a node promotes its children so a user can discard a
    heading without accidentally losing the BOQ rows beneath it.
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

    def rename_node(self, node_id: str, description: str) -> None:
        """Edit the description of any non-root tree node."""
        node = self.tree.find(node_id)
        if node is None or node.node_type == TreeNodeType.SHEET_ROOT:
            raise TreeEditError("The selected row cannot be renamed.")
        description = description.strip()
        if not description:
            raise TreeEditError("Description cannot be blank.")
        node.description = description

    def rename_wbs(self, node_id: str, description: str) -> None:
        """Rename only a WBS node; retained for backward compatibility."""
        node = self.tree.find(node_id)
        if node is None or not node.is_wbs:
            raise TreeEditError("Only WBS nodes can be renamed.")
        self.rename_node(node_id, description)

    def mark_as_wbs(self, node_id: str) -> None:
        """Convert a selected source row into a WBS heading."""
        location = self._required_location(node_id)
        node = location.parent.children[location.index]
        parent = location.parent
        if parent.node_type != TreeNodeType.SHEET_ROOT and not parent.is_wbs:
            raise TreeEditError("A WBS row must be placed under the sheet or another WBS row.")
        level = 2 + self.depth(node_id)
        if level not in {3, 4}:
            raise TreeEditError("Only WBS-3 and WBS-4 are supported.")
        node.node_type = TreeNodeType.WBS
        node.confidence = 1.0
        node.reason = "Marked as WBS by user"

    def mark_as_item(self, node_id: str) -> None:
        """Convert a selected WBS heading back into an ITEM row."""
        location = self._required_location(node_id)
        node = location.parent.children[location.index]
        node.node_type = TreeNodeType.ITEM
        node.confidence = 1.0
        node.reason = "Marked as ITEM by user"

    def toggle_wbs_item(self, node_ids: list[str]) -> TreeNodeType:
        """Toggle selected rows between WBS and ITEM.

        If every selected row is currently WBS, all become ITEM. Otherwise all
        selected rows become WBS. This makes mixed selections deterministic.
        """
        unique_ids = list(dict.fromkeys(node_ids))
        if not unique_ids:
            raise TreeEditError("Select at least one row.")
        nodes = []
        for node_id in unique_ids:
            node = self.tree.find(node_id)
            if node is None or node.node_type == TreeNodeType.SHEET_ROOT:
                raise TreeEditError("The selected row cannot be toggled.")
            nodes.append(node)

        target = TreeNodeType.ITEM if all(node.is_wbs for node in nodes) else TreeNodeType.WBS
        if target == TreeNodeType.WBS:
            # Validate every row first so the operation is atomic.
            for node_id in unique_ids:
                location = self._required_location(node_id)
                parent = location.parent
                if parent.node_type != TreeNodeType.SHEET_ROOT and not parent.is_wbs:
                    raise TreeEditError("A WBS row must be placed under the sheet or another WBS row.")
                level = 2 + self.depth(node_id)
                if level not in {3, 4}:
                    raise TreeEditError("Only WBS-3 and WBS-4 are supported.")
            for node_id in unique_ids:
                self.mark_as_wbs(node_id)
        else:
            for node_id in unique_ids:
                self.mark_as_item(node_id)
        return target

    def wbs_level(self, node_id: str) -> int | None:
        node = self.tree.find(node_id)
        if node is None or not node.is_wbs:
            return None
        level = 2 + self.depth(node_id)
        return level if level in {3, 4} else None

    def move_up(self, node_id: str) -> None:
        location = self._required_location(node_id)
        if location.index == 0:
            raise TreeEditError("The row is already first among its siblings.")
        siblings = location.parent.children
        siblings[location.index - 1], siblings[location.index] = (
            siblings[location.index],
            siblings[location.index - 1],
        )

    def move_down(self, node_id: str) -> None:
        location = self._required_location(node_id)
        siblings = location.parent.children
        if location.index >= len(siblings) - 1:
            raise TreeEditError("The row is already last among its siblings.")
        siblings[location.index], siblings[location.index + 1] = (
            siblings[location.index + 1],
            siblings[location.index],
        )


    def move_nodes_up(self, node_ids: list[str]) -> None:
        """Move selected sibling rows up one position while preserving order."""
        self._move_selected(node_ids, direction=-1)

    def move_nodes_down(self, node_ids: list[str]) -> None:
        """Move selected sibling rows down one position while preserving order."""
        self._move_selected(node_ids, direction=1)

    def delete_nodes(self, node_ids: list[str]) -> None:
        """Delete multiple rows safely, promoting each row's children."""
        unique_ids = list(dict.fromkeys(node_ids))
        # Delete descendants before ancestors and later siblings before earlier
        # siblings. This keeps all remaining node locations valid while editing.
        ordered = sorted(
            unique_ids,
            key=lambda node_id: (self.depth(node_id), self._required_location(node_id).index),
            reverse=True,
        )
        for node_id in ordered:
            if self.locate(node_id) is not None:
                self.delete_node(node_id)

    def _move_selected(self, node_ids: list[str], direction: int) -> None:
        if not node_ids:
            raise TreeEditError("Select at least one row.")
        groups: dict[str, tuple[DraftTreeNode, set[str]]] = {}
        for node_id in dict.fromkeys(node_ids):
            location = self._required_location(node_id)
            key = location.parent.node_id
            if key not in groups:
                groups[key] = (location.parent, set())
            groups[key][1].add(node_id)

        moved = False
        for parent, selected_ids in groups.values():
            siblings = parent.children
            selected_indexes = {i for i, node in enumerate(siblings) if node.node_id in selected_ids}
            if direction < 0:
                for index in range(1, len(siblings)):
                    if index in selected_indexes and index - 1 not in selected_indexes:
                        siblings[index - 1], siblings[index] = siblings[index], siblings[index - 1]
                        selected_indexes.remove(index)
                        selected_indexes.add(index - 1)
                        moved = True
            else:
                for index in range(len(siblings) - 2, -1, -1):
                    if index in selected_indexes and index + 1 not in selected_indexes:
                        siblings[index], siblings[index + 1] = siblings[index + 1], siblings[index]
                        selected_indexes.remove(index)
                        selected_indexes.add(index + 1)
                        moved = True
        if not moved:
            edge = "first" if direction < 0 else "last"
            raise TreeEditError(f"The selected rows are already at the {edge} available position.")

    def delete_node(self, node_id: str) -> None:
        location = self._required_location(node_id)
        node = location.parent.children[location.index]
        location.parent.children[location.index: location.index + 1] = node.children

    def delete_virtual_wbs(self, node_id: str) -> None:
        location = self._required_location(node_id)
        node = location.parent.children[location.index]
        if node.node_type != TreeNodeType.VIRTUAL_WBS:
            raise TreeEditError("Only a user-created WBS can be deleted.")
        self.delete_node(node_id)

    def _required_location(self, node_id: str) -> TreeLocation:
        location = self.locate(node_id)
        if location is None:
            raise TreeEditError("Tree node was not found.")
        return location
