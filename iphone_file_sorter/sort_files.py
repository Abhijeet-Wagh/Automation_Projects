#!/usr/bin/env python3
"""
Sort files from iPhone Internal Storage (or any folder tree) into
category folders by file extension.

Categories: Images, Videos, Documents, Excel, PDF, Other
"""

from __future__ import annotations

import argparse
import shutil
import sys
from collections import Counter
from pathlib import Path


CATEGORY_EXTENSIONS: dict[str, set[str]] = {
    "Images": {
        ".jpg",
        ".jpeg",
        ".png",
        ".heic",
        ".heif",
        ".gif",
        ".webp",
        ".bmp",
        ".tif",
        ".tiff",
        ".raw",
        ".dng",
    },
    "Videos": {
        ".mov",
        ".mp4",
        ".m4v",
        ".avi",
        ".mkv",
        ".3gp",
        ".mpg",
        ".mpeg",
    },
    "PDF": {".pdf"},
    "Excel": {".xls", ".xlsx", ".csv", ".xlsm", ".ods"},
    "Documents": {
        ".doc",
        ".docx",
        ".txt",
        ".rtf",
        ".pages",
        ".odt",
        ".ppt",
        ".pptx",
        ".key",
    },
}

# Build reverse lookup: extension -> category
EXTENSION_TO_CATEGORY: dict[str, str] = {
    ext: category
    for category, extensions in CATEGORY_EXTENSIONS.items()
    for ext in extensions
}

SKIP_NAMES = {".ds_store", "thumbs.db", "desktop.ini"}


def categorize(path: Path) -> str:
    """Return the category folder name for a file path."""
    ext = path.suffix.lower()
    return EXTENSION_TO_CATEGORY.get(ext, "Other")


def unique_destination(dest_dir: Path, filename: str) -> Path:
    """
    Return a destination path that does not overwrite an existing file.
    IMG_1.jpg -> IMG_1_1.jpg -> IMG_1_2.jpg ...
    """
    candidate = dest_dir / filename
    if not candidate.exists():
        return candidate

    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 1
    while True:
        candidate = dest_dir / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def iter_source_files(source: Path):
    """Yield regular files under source, skipping junk names."""
    for path in source.rglob("*"):
        if not path.is_file():
            continue
        if path.name.lower() in SKIP_NAMES:
            continue
        yield path


def sort_files(
    source: Path,
    destination: Path,
    *,
    dry_run: bool = False,
    move: bool = False,
) -> dict[str, int]:
    """
    Copy (or move) files from source into categorized folders under destination.

    Returns a counter of files handled per category.
    """
    if not source.exists() or not source.is_dir():
        raise FileNotFoundError(f"Source folder not found: {source}")

    counts: Counter[str] = Counter()
    action = "Moving" if move else "Copying"

    if not dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    for src_file in iter_source_files(source):
        # Do not re-process files already inside the destination tree
        try:
            src_file.resolve().relative_to(destination.resolve())
            continue
        except (ValueError, OSError):
            pass

        category = categorize(src_file)
        category_dir = destination / category
        dest_file = unique_destination(category_dir, src_file.name)

        print(f"{action}: {src_file} -> {dest_file}")
        counts[category] += 1

        if dry_run:
            continue

        category_dir.mkdir(parents=True, exist_ok=True)
        if move:
            shutil.move(str(src_file), str(dest_file))
        else:
            shutil.copy2(str(src_file), str(dest_file))

    return dict(counts)


def print_summary(counts: dict[str, int], dry_run: bool) -> None:
    total = sum(counts.values())
    mode = "DRY RUN — no files changed" if dry_run else "Done"
    print()
    print("=" * 40)
    print(mode)
    print("=" * 40)
    if not counts:
        print("No files found.")
        return

    for category in ("Images", "Videos", "Documents", "Excel", "PDF", "Other"):
        if category in counts:
            print(f"  {category:12} {counts[category]}")
    # Any unexpected keys
    for category, count in sorted(counts.items()):
        if category not in {"Images", "Videos", "Documents", "Excel", "PDF", "Other"}:
            print(f"  {category:12} {count}")
    print(f"  {'Total':12} {total}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Sort iPhone Internal Storage files (or any folder tree) "
            "into Images, Videos, Documents, Excel, PDF, and Other."
        )
    )
    parser.add_argument(
        "source",
        type=Path,
        help="Source folder (e.g. copied Internal Storage directory)",
    )
    parser.add_argument(
        "destination",
        type=Path,
        help="Destination folder for sorted categories",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would happen without copying or moving files",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help="Move files instead of copying (default is copy)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    source = args.source.expanduser().resolve()
    destination = args.destination.expanduser().resolve()

    if source == destination:
        print("Error: source and destination must be different folders.", file=sys.stderr)
        return 1

    try:
        destination.relative_to(source)
        print(
            "Error: destination cannot be inside the source folder.",
            file=sys.stderr,
        )
        return 1
    except ValueError:
        pass

    try:
        counts = sort_files(
            source,
            destination,
            dry_run=args.dry_run,
            move=args.move,
        )
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print_summary(counts, args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
