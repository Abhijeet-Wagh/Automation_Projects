"""Folder tree helpers for checkbox selection with parent/child cascade."""

from __future__ import annotations

from typing import Any


TreeNode = dict[str, Any]


def iter_nodes(node: TreeNode):
    """Yield all nodes in pre-order."""
    yield node
    for child in node.get("children") or []:
        yield from iter_nodes(child)


def all_paths(node: TreeNode) -> list[str]:
    return [n["path"] for n in iter_nodes(node)]


def descendant_paths(node: TreeNode) -> list[str]:
    """Node path + all descendant paths."""
    return all_paths(node)


def find_node(root: TreeNode, path: str) -> TreeNode | None:
    for node in iter_nodes(root):
        if node.get("path") == path:
            return node
    return None


def parent_path(path: str) -> str | None:
    if path == "":
        return None
    if "/" not in path:
        return ""
    return path.rsplit("/", 1)[0]


def apply_check(
    selected: set[str],
    root: TreeNode,
    path: str,
    checked: bool,
) -> set[str]:
    """
    Toggle a node and cascade.

    - Checking a folder selects it and all subfolders.
    - Unchecking a folder deselects it and all subfolders.
    - After a child change, ancestors are checked only if all their
      descendants are selected; otherwise ancestors are unchecked.
    """
    node = find_node(root, path)
    if node is None:
        return set(selected)

    result = set(selected)
    affected = descendant_paths(node)
    if checked:
        result.update(affected)
    else:
        result.difference_update(affected)

    # Update ancestors based on whether all children are selected
    current = parent_path(path)
    while current is not None:
        ancestor = find_node(root, current)
        if ancestor is None:
            break
        child_paths = [c["path"] for c in ancestor.get("children") or []]
        if child_paths and all(p in result for p in child_paths):
            result.add(current)
        else:
            result.discard(current)
        current = parent_path(current)

    return result


def selection_roots(selected: set[str]) -> list[str]:
    """
    Minimal folders to copy: selected nodes whose parent is not selected.
    If root "" is selected, returns [""].
    """
    if "" in selected:
        return [""]

    roots: list[str] = []
    for path in sorted(selected):
        parent = parent_path(path)
        covered = False
        while parent is not None:
            if parent in selected:
                covered = True
                break
            parent = parent_path(parent)
        if not covered:
            roots.append(path)
    return roots


def count_folders(node: TreeNode) -> int:
    """Count folder nodes under node, excluding the node itself."""
    return max(0, len(list(iter_nodes(node))) - 1)
