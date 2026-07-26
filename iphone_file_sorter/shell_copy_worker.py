#!/usr/bin/env python3
"""
Windows STA worker: copy one iPhone folder via Shell.Application.

Streamlit threads do not pump Win32 messages, so in-process CopyHere often
never actually transfers files. This script runs as its own process, pumps
messages, and exits when the destination folder has a stable file count.

Usage:
  python shell_copy_worker.py <device_name> <folder_rel> <destination_dir> [timeout_sec]
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

THIS_PC = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
FOF_NOCONFIRMATION = 16
FOF_NOCONFIRMMKDIR = 512
COPY_FLAGS = FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR


def count_files(folder: Path) -> int:
    if not folder.exists():
        return 0
    return sum(1 for p in folder.rglob("*") if p.is_file())


def find_device_folder(shell, device_name: str):
    computer = shell.NameSpace(THIS_PC)
    if computer is None:
        raise RuntimeError("Could not open This PC")
    for item in computer.Items():
        if str(item.Name) == device_name:
            folder = item.GetFolder
            if folder is None:
                raise RuntimeError(f"Could not open device: {device_name}")
            return folder
    raise RuntimeError(f"Device not found: {device_name}")


def find_internal_storage(device_folder):
    names = [str(i.Name) for i in device_folder.Items()]
    for candidate in ("Internal Storage", "Internal storage", "DCIM"):
        if candidate in names:
            for item in device_folder.Items():
                if str(item.Name) == candidate:
                    return item.GetFolder, candidate
    return device_folder, str(getattr(device_folder, "Title", "device"))


def get_child_folder(parent, child_name: str):
    for item in parent.Items():
        if str(item.Name) == child_name:
            folder = item.GetFolder
            if folder is None:
                raise RuntimeError(f"Not a folder: {child_name}")
            return folder
    raise RuntimeError(f"Folder not found: {child_name}")


def resolve_folder_item(media_root, folder_rel: str):
    parts = [p for p in folder_rel.replace("\\", "/").split("/") if p]
    if not parts:
        raise RuntimeError("Empty folder path")
    parent = media_root
    for part in parts[:-1]:
        parent = get_child_folder(parent, part)
    name = parts[-1]
    for item in parent.Items():
        if str(item.Name) == name:
            return item
    raise RuntimeError(f"Folder item not found: {folder_rel}")


def pump():
    import pythoncom  # type: ignore

    pythoncom.PumpWaitingMessages()


def main(argv: list[str]) -> int:
    if len(argv) < 4:
        print(
            "Usage: shell_copy_worker.py <device_name> <folder_rel> "
            "<destination_dir> [timeout_sec]",
            file=sys.stderr,
        )
        return 2

    device_name = argv[1]
    folder_rel = argv[2]
    destination = Path(argv[3])
    timeout_sec = float(argv[4]) if len(argv) > 4 else 7200.0

    destination.mkdir(parents=True, exist_ok=True)
    dest_folder = destination / folder_rel

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        device = find_device_folder(shell, device_name)
        media_root, _label = find_internal_storage(device)
        item = resolve_folder_item(media_root, folder_rel)

        dest_ns = shell.NameSpace(str(destination.resolve()))
        if dest_ns is None:
            raise RuntimeError(f"Could not open destination: {destination}")

        print(f"COPY_START {folder_rel}", flush=True)
        dest_ns.CopyHere(item, COPY_FLAGS)

        start = time.time()
        last_count = -1
        stable = 0
        while time.time() - start < timeout_sec:
            pump()
            count = count_files(dest_folder)
            elapsed = int(time.time() - start)
            if elapsed % 2 == 0:
                print(f"COPY_PROGRESS files={count} elapsed={elapsed}", flush=True)

            if count > 0 and count == last_count:
                stable += 1
                # 20 * 0.1s ~= 2 seconds with no new files
                if stable >= 20:
                    print(f"COPY_DONE files={count}", flush=True)
                    return 0
            else:
                stable = 0
                last_count = count
            time.sleep(0.1)

        count = count_files(dest_folder)
        if count > 0:
            print(f"COPY_DONE files={count} (timeout with files)", flush=True)
            return 0
        print("COPY_FAIL no files appeared at destination", flush=True)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"COPY_FAIL {exc}", flush=True)
        return 1
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
