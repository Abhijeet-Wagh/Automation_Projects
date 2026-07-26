"""
Curated iPhone content categories for copy.

AFC (USB file access) can pull camera roll, recordings, downloads, and sometimes
app media. SMS / Contacts / WhatsApp chat databases need a selective iPhone backup.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


SourceKind = Literal["afc_media", "afc_app", "backup"]


IMAGE_EXTS = {
    ".heic",
    ".heif",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".tif",
    ".tiff",
    ".bmp",
    ".webp",
    ".dng",
    ".raw",
}
VIDEO_EXTS = {
    ".mov",
    ".mp4",
    ".m4v",
    ".3gp",
    ".avi",
    ".mkv",
    ".mpg",
    ".mpeg",
}
AUDIO_EXTS = {
    ".m4a",
    ".mp3",
    ".aac",
    ".caf",
    ".wav",
    ".aiff",
    ".aif",
    ".flac",
    ".ogg",
    ".amr",
}
DOC_EXTS = {
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".csv",
    ".ppt",
    ".pptx",
    ".txt",
    ".rtf",
    ".pages",
    ".numbers",
    ".key",
    ".odt",
    ".ods",
    ".odp",
    ".epub",
    ".html",
    ".htm",
}


@dataclass(frozen=True)
class ContentCategory:
    id: str
    label: str
    description: str
    kind: SourceKind
    # AFC media path under /var/mobile/Media (e.g. DCIM), ignored for backup/app
    media_remote: str = ""
    # App bundle for house_arrest
    bundle_id: str = ""
    # Allowed file suffixes (lowercase, with dot). Empty = keep filter via skip-rules only.
    extensions: frozenset[str] = frozenset()
    # pymobiledevice3 BackupSelection names
    backup_selections: tuple[str, ...] = ()
    # Extra backup path regexes (notes, etc.)
    backup_regexes: tuple[str, ...] = ()
    default_on: bool = False
    group: str = "media"  # media | messages | people


CATEGORIES: list[ContentCategory] = [
    ContentCategory(
        id="camera",
        label="Camera photos & videos",
        description="DCIM albums (100APPLE, 127APPLE, …) — your Camera Roll originals.",
        kind="afc_media",
        media_remote="DCIM",
        extensions=frozenset(IMAGE_EXTS | VIDEO_EXTS),
        default_on=True,
        group="media",
    ),
    ContentCategory(
        id="recordings",
        label="Voice memos & audio recordings",
        description="Built-in Voice Memos / Recordings folder.",
        kind="afc_media",
        media_remote="Recordings",
        extensions=frozenset(AUDIO_EXTS | VIDEO_EXTS),
        default_on=True,
        group="media",
    ),
    ContentCategory(
        id="downloads",
        label="Downloads (PDFs, documents)",
        description="Files saved to the iPhone Downloads area.",
        kind="afc_media",
        media_remote="Downloads",
        extensions=frozenset(DOC_EXTS | IMAGE_EXTS | AUDIO_EXTS | VIDEO_EXTS),
        default_on=True,
        group="media",
    ),
    ContentCategory(
        id="whatsapp_media",
        label="WhatsApp media (images / video / audio)",
        description=(
            "Tries USB access to the WhatsApp app container and copies media files only. "
            "If iOS blocks this, use WhatsApp Export or the backup option for chat DB."
        ),
        kind="afc_app",
        bundle_id="net.whatsapp.WhatsApp",
        extensions=frozenset(IMAGE_EXTS | VIDEO_EXTS | AUDIO_EXTS | DOC_EXTS),
        default_on=True,
        group="media",
    ),
    ContentCategory(
        id="telegram_media",
        label="Telegram media",
        description="Tries USB access to Telegram; copies media/docs only when allowed.",
        kind="afc_app",
        bundle_id="ph.telegra.Telegraph",
        extensions=frozenset(IMAGE_EXTS | VIDEO_EXTS | AUDIO_EXTS | DOC_EXTS),
        default_on=False,
        group="media",
    ),
    ContentCategory(
        id="sms",
        label="SMS / iMessage database",
        description="Selective iPhone backup of sms.db (not readable as plain text here).",
        kind="backup",
        backup_selections=("sms",),
        default_on=False,
        group="messages",
    ),
    ContentCategory(
        id="whatsapp_chats",
        label="WhatsApp chat database",
        description=(
            "Selective backup of ChatStorage.sqlite. "
            "Open later with a WhatsApp iOS backup viewer/exporter."
        ),
        kind="backup",
        backup_selections=("whatsapp",),
        default_on=False,
        group="messages",
    ),
    ContentCategory(
        id="contacts",
        label="Contacts",
        description="Selective backup of AddressBook.sqlitedb.",
        kind="backup",
        backup_selections=("contacts",),
        default_on=False,
        group="people",
    ),
    ContentCategory(
        id="call_history",
        label="Call history",
        description="Selective backup of CallHistory store.",
        kind="backup",
        backup_selections=("call_history",),
        default_on=False,
        group="people",
    ),
    ContentCategory(
        id="notes",
        label="Apple Notes (backup files)",
        description=(
            "Selective backup of Notes-related files (best-effort path match). "
            "Can make backup much slower — try Contacts alone first if a run hangs."
        ),
        kind="backup",
        backup_regexes=(
            r"(?i)Notes",
            r"(?i)group\.com\.apple\.notes",
        ),
        default_on=False,
        group="people",
    ),
]


CATEGORY_BY_ID = {c.id: c for c in CATEGORIES}


def categories_by_group() -> dict[str, list[ContentCategory]]:
    groups: dict[str, list[ContentCategory]] = {
        "media": [],
        "messages": [],
        "people": [],
    }
    for cat in CATEGORIES:
        groups.setdefault(cat.group, []).append(cat)
    return groups


def default_selected_ids() -> list[str]:
    return [c.id for c in CATEGORIES if c.default_on]
