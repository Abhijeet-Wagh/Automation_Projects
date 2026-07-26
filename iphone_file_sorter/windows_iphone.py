"""
Detect a connected iPhone/iPad in Windows Explorer (This PC) and copy
files using the Windows Shell — the same engine as File Explorer.

Copies file-by-file, skips failures, and returns detailed results for Excel logging.

Requires: pywin32 (Windows only)
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from copy_log import default_log_path, write_copy_excel_log

THIS_PC = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"

# Shell CopyHere flags
FOF_SILENT = 4
FOF_NOCONFIRMATION = 16
FOF_NOCONFIRMMKDIR = 512
FOF_NOERRORUI = 1024

# Quiet auto-skip style flags (no confirm / minimal UI)
COPY_FLAGS_SKIP = (
    FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR | FOF_SILENT | FOF_NOERRORUI
)


class WindowsShellError(RuntimeError):
    pass


@dataclass
class CopyResult:
    succeeded: list[dict[str, Any]] = field(default_factory=list)
    failed: list[dict[str, Any]] = field(default_factory=list)
    log_path: Path | None = None
    selected_folders: list[str] = field(default_factory=list)


def _shell():
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise WindowsShellError(
            "pywin32 is required on Windows. Run: python -m pip install pywin32"
        ) from exc
    return win32com.client.Dispatch("Shell.Application")


def is_windows() -> bool:
    import sys

    return sys.platform.startswith("win")


def list_portable_apple_devices() -> list[str]:
    """Return device names under This PC that look like iPhone/iPad."""
    if not is_windows():
        return []

    shell = _shell()
    computer = shell.NameSpace(THIS_PC)
    if computer is None:
        return []

    names: list[str] = []
    for item in computer.Items():
        name = str(item.Name)
        lower = name.lower()
        if "iphone" in lower or "ipad" in lower or "apple" in lower:
            names.append(name)
    return names


def _device_folder(device_name: str) -> Any:
    shell = _shell()
    computer = shell.NameSpace(THIS_PC)
    if computer is None:
        raise WindowsShellError("Could not open This PC.")

    for item in computer.Items():
        if str(item.Name) == device_name:
            folder = item.GetFolder
            if folder is None:
                raise WindowsShellError(f"Could not open device folder: {device_name}")
            return folder
    raise WindowsShellError(f"Device not found under This PC: {device_name}")


def list_child_names(folder: Any) -> list[str]:
    return [str(item.Name) for item in folder.Items()]


def get_child_folder(parent_folder: Any, child_name: str) -> Any:
    for item in parent_folder.Items():
        if str(item.Name) == child_name:
            folder = item.GetFolder
            if folder is None:
                raise WindowsShellError(f"Not a folder: {child_name}")
            return folder
    raise WindowsShellError(f"Folder not found: {child_name}")


def find_internal_storage(device_name: str) -> tuple[Any, str]:
    """
    Open the device and locate Internal Storage (or the device root if
    media folders are directly underneath).
    """
    device = _device_folder(device_name)
    children = list_child_names(device)

    for candidate in ("Internal Storage", "Internal storage", "DCIM"):
        if candidate in children:
            return get_child_folder(device, candidate), candidate

    return device, device_name


def list_iphone_media_folders(device_name: str) -> list[str]:
    """List copyable folders under the iPhone media root."""
    root, _ = find_internal_storage(device_name)
    return sorted(list_child_names(root))


def _is_folder_item(item: Any) -> bool:
    try:
        return bool(item.IsFolder)
    except Exception:
        try:
            return item.GetFolder is not None and not hasattr(item, "Size")
        except Exception:
            return False


def _enumerate_relative_files(folder: Any, prefix: str = "") -> list[str]:
    """Return relative file paths under a Shell folder (files only)."""
    paths: list[str] = []
    for item in list(folder.Items()):
        name = str(item.Name)
        rel = f"{prefix}/{name}" if prefix else name
        if _is_folder_item(item):
            child = item.GetFolder
            if child is not None:
                paths.extend(_enumerate_relative_files(child, rel))
        else:
            paths.append(rel)
    return paths


def _resolve_file_item(
    media_root: Any, top_folder: str, relative_path: str
) -> tuple[Any, Any]:
    """Return (file_item, parent_folder) for a relative path under top_folder."""
    folder = get_child_folder(media_root, top_folder)
    parts = relative_path.replace("\\", "/").split("/")
    for part in parts[:-1]:
        folder = get_child_folder(folder, part)

    filename = parts[-1]
    for item in folder.Items():
        if str(item.Name) == filename and not _is_folder_item(item):
            return item, folder
    raise WindowsShellError(f"File not found on iPhone: {top_folder}/{relative_path}")


def _wait_for_file(path: Path, timeout_sec: float = 90.0) -> tuple[bool, str]:
    """
    Wait until a copied file exists with a stable non-zero size.
    Returns (ok, reason_if_not_ok).
    """
    start = time.time()
    last_size = -1
    stable_rounds = 0
    saw_exists = False

    while time.time() - start < timeout_sec:
        if path.exists() and path.is_file():
            saw_exists = True
            try:
                size = path.stat().st_size
            except OSError:
                time.sleep(0.4)
                continue

            if size > 0 and size == last_size:
                stable_rounds += 1
                if stable_rounds >= 2:
                    return True, ""
            else:
                stable_rounds = 0
                last_size = size
        time.sleep(0.4)

    if saw_exists:
        try:
            size = path.stat().st_size if path.exists() else 0
        except OSError:
            size = 0
        if size <= 0:
            return False, "File appeared but stayed empty (likely MTP/copy failure)"
        return False, "Timed out waiting for file size to stabilize"
    return False, "Timed out — file did not appear at destination"


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def copy_folders_file_by_file(
    device_name: str,
    folder_names: list[str],
    destination: Path,
    *,
    log_path: Path | None = None,
    file_timeout_sec: float = 90.0,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> CopyResult:
    """
    Copy all files under the selected iPhone folders to destination.

    - Works file-by-file
    - On error/timeout for one file: log it, skip, continue
    - Writes an Excel log with failed/successful details
    """
    if not folder_names:
        return CopyResult(selected_folders=[])

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    if log_path is None:
        log_path = default_log_path(destination)

    shell = _shell()
    media_root, media_label = find_internal_storage(device_name)

    # Build job list: (top_folder, relative_path)
    normalized_jobs: list[tuple[str, str]] = []
    failed: list[dict[str, Any]] = []
    for folder_name in folder_names:
        try:
            folder = get_child_folder(media_root, folder_name)
        except WindowsShellError as exc:
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": folder_name,
                    "file_name": "",
                    "relative_path": "",
                    "source_path": f"{device_name}/{media_label}/{folder_name}",
                    "destination_path": str(destination / folder_name),
                    "error": str(exc),
                    "status": "Skipped",
                }
            )
            continue

        for rel in _enumerate_relative_files(folder):
            normalized_jobs.append((folder_name, rel))

    succeeded: list[dict[str, Any]] = []
    total = len(normalized_jobs)

    for index, (folder_name, rel) in enumerate(normalized_jobs, start=1):
        file_name = Path(rel).name
        source_path = f"{device_name}/{media_label}/{folder_name}/{rel}"
        dest_file = destination / folder_name / Path(rel)
        dest_file.parent.mkdir(parents=True, exist_ok=True)

        label = f"{folder_name}/{rel}"
        if on_progress:
            on_progress(label, index, total, "copying")

        try:
            item, _parent = _resolve_file_item(media_root, folder_name, rel)
            dest_ns = shell.NameSpace(str(dest_file.parent.resolve()))
            if dest_ns is None:
                raise WindowsShellError(
                    f"Could not open destination folder: {dest_file.parent}"
                )

            # Remove incomplete leftover from a previous failed attempt
            if dest_file.exists():
                try:
                    if dest_file.stat().st_size == 0:
                        dest_file.unlink()
                except OSError:
                    pass

            dest_ns.CopyHere(item, COPY_FLAGS_SKIP)
            ok, reason = _wait_for_file(dest_file, timeout_sec=file_timeout_sec)
            if not ok:
                # Clean empty stub if present
                try:
                    if dest_file.exists() and dest_file.stat().st_size == 0:
                        dest_file.unlink()
                except OSError:
                    pass
                raise WindowsShellError(reason or "Copy failed")

            succeeded.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": folder_name,
                    "file_name": file_name,
                    "relative_path": rel,
                    "destination_path": str(dest_file),
                    "status": "Copied",
                }
            )
            if on_progress:
                on_progress(label, index, total, "done")

        except Exception as exc:  # noqa: BLE001 - skip and continue
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": folder_name,
                    "file_name": file_name,
                    "relative_path": rel,
                    "source_path": source_path,
                    "destination_path": str(dest_file),
                    "error": str(exc),
                    "status": "Skipped",
                }
            )
            if on_progress:
                on_progress(label, index, total, "skipped")
            continue

    write_copy_excel_log(
        log_path,
        device_name=device_name,
        selected_folders=folder_names,
        succeeded=succeeded,
        failed=failed,
    )

    return CopyResult(
        succeeded=succeeded,
        failed=failed,
        log_path=Path(log_path),
        selected_folders=list(folder_names),
    )


# Backward-compatible wrapper used by older UI code paths
def copy_named_items_to_folder(
    device_name: str,
    item_names: list[str],
    destination: Path,
    *,
    on_progress=None,
    silent: bool = True,
    log_path: Path | None = None,
    file_timeout_sec: float = 90.0,
) -> CopyResult:
    """Copy selected folders file-by-file with skip-on-error + Excel log."""
    del silent  # always skip quietly at file level
    return copy_folders_file_by_file(
        device_name,
        item_names,
        destination,
        log_path=log_path,
        file_timeout_sec=file_timeout_sec,
        on_progress=on_progress,
    )
