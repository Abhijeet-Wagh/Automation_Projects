#!/usr/bin/env python3
"""Tests for folder tree selection cascade."""

from __future__ import annotations

import unittest

from folder_tree import apply_check, selection_roots


def sample_tree():
    return {
        "name": "Internal Storage",
        "path": "",
        "children": [
            {
                "name": "202403_a",
                "path": "202403_a",
                "children": [
                    {"name": "sub1", "path": "202403_a/sub1", "children": []},
                    {"name": "sub2", "path": "202403_a/sub2", "children": []},
                ],
            },
            {"name": "202403_b", "path": "202403_b", "children": []},
        ],
    }


class FolderTreeTests(unittest.TestCase):
    def test_check_parent_selects_children(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "202403_a", True)
        self.assertIn("202403_a", selected)
        self.assertIn("202403_a/sub1", selected)
        self.assertIn("202403_a/sub2", selected)
        self.assertNotIn("202403_b", selected)

    def test_uncheck_parent_clears_children(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "202403_a", True)
        selected = apply_check(selected, tree, "202403_a", False)
        self.assertNotIn("202403_a", selected)
        self.assertNotIn("202403_a/sub1", selected)

    def test_uncheck_child_unchecks_parent(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "202403_a", True)
        selected = apply_check(selected, tree, "202403_a/sub1", False)
        self.assertNotIn("202403_a", selected)
        self.assertNotIn("202403_a/sub1", selected)
        self.assertIn("202403_a/sub2", selected)

    def test_check_all_children_checks_parent(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "202403_a/sub1", True)
        selected = apply_check(selected, tree, "202403_a/sub2", True)
        self.assertIn("202403_a", selected)

    def test_selection_roots_skips_covered_children(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "202403_a", True)
        roots = selection_roots(selected)
        self.assertEqual(roots, ["202403_a"])

    def test_selection_roots_full_tree(self):
        tree = sample_tree()
        selected = apply_check(set(), tree, "", True)
        self.assertEqual(selection_roots(selected), [""])


if __name__ == "__main__":
    unittest.main()
