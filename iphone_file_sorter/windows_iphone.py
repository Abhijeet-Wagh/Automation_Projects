"""
Detect a connected iPhone/iPad in Windows Explorer (This PC) and copy
folders/files using the Windows Shell — the same engine as File Explorer.

Requires: pywin32 (Windows only)
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

THIS_PC = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"

# Shell CopyHere flags
FOF_SILENT = 4
FOF_NOCONFIRMATION = 16
FOF_NOCONFIRMMKDIR = 512
FOF_NOERRORUI = 1024


class WindowsShellError(RuntimeError):
    pass


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

    # Some iPhones expose dated folders directly under the device
    return device, device_name


def list_iphone_media_folders(device_name: str) -> list[str]:
    """List copyable folders under the iPhone media root."""
    root, _ = find_internal_storage(device_name)
    return sorted(list_child_names(root))


def _wait_for_destination_item(dest: Path, name: str, timeout_sec: int = 7200) -> bool:
    """Wait until a copied item appears in the destination (Shell CopyHere is async)."""
    target = dest / name
    start = time.time()
    last_size = -1
    stable_rounds = 0

    while time.time() - start < timeout_sec:
        if target.exists():
            try:
                if target.is_file():
                    size = target.stat().st_size
                    if size == last_size and size > 0:
                        stable_rounds += 1
                        if stable_rounds >= 2:
                            return True
                    else:
                        stable_rounds = 0
                        last_size = size
                else:
                    # Folder: consider present once it exists and is non-empty
                    # or after a short settle period.
                    if any(target.iterdir()) or (time.time() - start) > 2:
                        time.sleep(1.0)
                        return True
            except OSError:
                pass
        time.sleep(0.5)
    return target.exists()


def copy_named_items_to_folder(
    device_name: str,
    item_names: list[str],
    destination: Path,
    *,
    on_progress=None,
    silent: bool = False,
) -> list[str]:
    """
    Copy selected folders/files from the iPhone media root to destination.

    Uses Windows Shell CopyHere (same as Explorer). Returns list of copied names.
    """
    if not item_names:
        return []

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)

    shell = _shell()
    media_root, _ = find_internal_storage(device_name)
    dest_ns = shell.NameSpace(str(destination.resolve()))
    if dest_ns is None:
        raise WindowsShellError(f"Could not open destination: {destination}")

    flags = FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR
    if silent:
        flags |= FOF_SILENT | FOF_NOERRORUI

    copied: list[str] = []
    total = len(item_names)

    for index, name in enumerate(item_names, start=1):
        item_obj = None
        for item in media_root.Items():
            if str(item.Name) == name:
                item_obj = item
                break
        if item_obj is None:
            raise WindowsShellError(f"Item not found on iPhone: {name}")

        if on_progress:
            on_progress(name, index, total, "starting")

        dest_ns.CopyHere(item_obj, flags)
        ok = _wait_for_destination_item(destination, name)
        if not ok:
            raise WindowsShellError(
                f"Timed out waiting for copy of '{name}'. "
                "Check the Windows copy dialog for errors, keep the iPhone unlocked, "
                "and try copying fewer folders at a time."
            )

        copied.append(name)
        if on_progress:
            on_progress(name, index, total, "done")

    return copied
