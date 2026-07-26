"""Tests for curated content categories."""

from __future__ import annotations

from content_categories import (
    CATEGORIES,
    CATEGORY_BY_ID,
    default_selected_ids,
    categories_by_group,
)
from afc_iphone import _ext_allowed, _should_skip_remote


def test_default_categories_are_media_focused():
    ids = default_selected_ids()
    assert "camera" in ids
    assert "whatsapp_media" in ids
    assert "photodata" not in ids


def test_category_lookup():
    assert CATEGORY_BY_ID["camera"].media_remote == "DCIM"
    assert ".heic" in CATEGORY_BY_ID["camera"].extensions
    assert ".pdf" in CATEGORY_BY_ID["downloads"].extensions


def test_groups_cover_all_categories():
    grouped = categories_by_group()
    flat = [c.id for cats in grouped.values() for c in cats]
    assert sorted(flat) == sorted(c.id for c in CATEGORIES)


def test_extension_filter_keeps_camera_media():
    exts = set(CATEGORY_BY_ID["camera"].extensions)
    assert _ext_allowed("DCIM/127APPLE/IMG_1.HEIC", exts)
    assert _ext_allowed("DCIM/127APPLE/VID.MOV", exts)
    assert not _ext_allowed("DCIM/127APPLE/Thumbs.db", exts)
    assert _should_skip_remote("PhotoData/Photos.sqlite")
