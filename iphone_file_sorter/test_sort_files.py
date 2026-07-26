#!/usr/bin/env python3
"""Tests for sort_files.py"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from sort_files import categorize, sort_files, unique_destination


class CategorizeTests(unittest.TestCase):
    def test_images(self):
        self.assertEqual(categorize(Path("a.HEIC")), "Images")
        self.assertEqual(categorize(Path("b.jpg")), "Images")

    def test_videos(self):
        self.assertEqual(categorize(Path("clip.MOV")), "Videos")

    def test_pdf_excel_docs(self):
        self.assertEqual(categorize(Path("x.pdf")), "PDF")
        self.assertEqual(categorize(Path("sheet.xlsx")), "Excel")
        self.assertEqual(categorize(Path("note.docx")), "Documents")

    def test_other(self):
        self.assertEqual(categorize(Path("app.ipa")), "Other")


class UniqueDestinationTests(unittest.TestCase):
    def test_adds_suffix_when_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp)
            (dest / "photo.jpg").write_bytes(b"1")
            result = unique_destination(dest, "photo.jpg")
            self.assertEqual(result.name, "photo_1.jpg")


class SortFilesTests(unittest.TestCase):
    def test_copy_into_categories(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source"
            dest = root / "sorted"
            nested = source / "202403_b"
            nested.mkdir(parents=True)
            (nested / "pic.heic").write_bytes(b"img")
            (nested / "clip.mov").write_bytes(b"vid")
            (nested / "report.pdf").write_bytes(b"pdf")
            (nested / "data.xlsx").write_bytes(b"xls")
            (nested / "notes.docx").write_bytes(b"doc")
            (nested / "weird.xyz").write_bytes(b"x")

            counts = sort_files(source, dest, dry_run=False, move=False)

            self.assertEqual(counts["Images"], 1)
            self.assertEqual(counts["Videos"], 1)
            self.assertEqual(counts["PDF"], 1)
            self.assertEqual(counts["Excel"], 1)
            self.assertEqual(counts["Documents"], 1)
            self.assertEqual(counts["Other"], 1)
            self.assertTrue((dest / "Images" / "pic.heic").exists())
            self.assertTrue((dest / "Videos" / "clip.mov").exists())
            # Originals remain when copying
            self.assertTrue((nested / "pic.heic").exists())

    def test_dry_run_does_not_write(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source"
            dest = root / "sorted"
            source.mkdir()
            (source / "a.jpg").write_bytes(b"img")

            counts = sort_files(source, dest, dry_run=True, move=False)

            self.assertEqual(counts["Images"], 1)
            self.assertFalse(dest.exists())


if __name__ == "__main__":
    unittest.main()
