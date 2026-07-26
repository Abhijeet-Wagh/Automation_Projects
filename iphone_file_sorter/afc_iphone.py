"""
Copy real files from iPhone using Apple AFC (via pymobiledevice3).

This bypasses Windows MTP, which often creates empty folders and never
transfers file contents.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from copy_log import default_log_path, write_copy_excel_log


class AfcError(RuntimeError):
    pass


@dataclass
class AfcCopyResult:
    succeeded: list[dict[str, Any]] = field(default_factory=list)
    failed: list[dict[str, Any]] = field(default_factory=list)
    log_path: Path | None = None
    selected_folders: list[str] = field(default_factory=list)


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _run(coro):
    try:
        return asyncio.run(coro)
    except RuntimeError:
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(coro)
        finally:
            loop.close()


def afc_available() -> bool:
    try:
        import pymobiledevice3  # noqa: F401

        return True
    except ImportError:
        return False


async def _listdir(afc, path: str) -> list[str]:
    target = path if path else "."
    try:
        entries = await afc.listdir(target)
    except Exception:
        return []
    out: list[str] = []
    for name in entries:
        name = str(name)
        if name in {".", ".."}:
            continue
        out.append(name)
    return sorted(out, key=str.lower)


async def _isdir(afc, path: str) -> bool:
    target = path if path else "."
    try:
        return bool(await afc.isdir(target))
    except Exception:
        return False


def _join(parent: str, child: str) -> str:
    if not parent:
        return child
    return f"{parent}/{child}"


async def _list_devices_async() -> list[dict[str, str]]:
    from pymobiledevice3 import usbmux
    from pymobiledevice3.lockdown import create_using_usbmux

    devices = await usbmux.list_devices()
    result: list[dict[str, str]] = []
    for device in devices:
        connection_type = getattr(device, "connection_type", None)
        if connection_type and str(connection_type).upper() != "USB":
            continue
        serial = str(device.serial)
        name = serial
        try:
            async with await create_using_usbmux(serial=serial) as lockdown:
                name = str(getattr(lockdown, "display_name", None) or serial)
        except Exception:
            name = serial
        result.append({"serial": serial, "name": name})
    return result


def list_afc_devices() -> list[dict[str, str]]:
    """Return [{'serial': ..., 'name': ...}, ...] for USB iPhones."""
    if not afc_available():
        raise AfcError(
            "pymobiledevice3 is not installed. Run: python -m pip install pymobiledevice3"
        )
    try:
        return _run(_list_devices_async())
    except Exception as exc:  # noqa: BLE001
        raise AfcError(
            f"Could not list iPhone via Apple drivers: {exc}. "
            "Unlock the iPhone, tap Trust, and ensure iTunes/Apple Devices is installed."
        ) from exc


async def _build_tree_async(serial: str, max_depth: int = 2) -> dict:
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.afc import AfcService

    async with await create_using_usbmux(serial=serial) as lockdown:
        async with AfcService(lockdown=lockdown) as afc:
            root_path = "DCIM" if await _isdir(afc, "DCIM") else ""
            root_name = "DCIM" if root_path == "DCIM" else "Media"

            async def build(path: str, name: str, depth: int) -> dict:
                node: dict[str, Any] = {"name": name, "path": path, "children": []}
                if depth >= max_depth:
                    return node
                for child_name in await _listdir(afc, path):
                    child_path = _join(path, child_name)
                    if await _isdir(afc, child_path):
                        node["children"].append(
                            await build(child_path, child_name, depth + 1)
                        )
                return node

            return await build(root_path, root_name, 0)


def list_afc_folder_tree(serial: str, *, max_depth: int = 2) -> dict:
    try:
        return _run(_build_tree_async(serial, max_depth=max_depth))
    except AfcError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AfcError(f"Failed to read iPhone folders via AFC: {exc}") from exc


async def _copy_folders_async(
    serial: str,
    folder_paths: list[str],
    destination: Path,
    *,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> AfcCopyResult:
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.afc import AfcService

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    log_path = default_log_path(destination)

    succeeded: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []
    total = max(len(folder_paths), 1)
    device_name = serial

    async with await create_using_usbmux(serial=serial) as lockdown:
        device_name = str(getattr(lockdown, "display_name", None) or serial)
        async with AfcService(lockdown=lockdown) as afc:
            for index, folder_path in enumerate(folder_paths, start=1):
                remote = folder_path if folder_path else "DCIM"
                if on_progress:
                    on_progress(remote, index, total, "copying")

                before = {
                    str(p.relative_to(destination))
                    for p in destination.rglob("*")
                    if p.is_file()
                }

                try:
                    file_counter = {"n": 0}

                    def _cb(_src, _dst, _folder=remote):
                        file_counter["n"] += 1
                        if on_progress:
                            on_progress(
                                f"{_folder} ({file_counter['n']} files)",
                                index,
                                total,
                                "copying",
                            )

                    await afc.pull(
                        remote,
                        str(destination),
                        ignore_errors=True,
                        progress_bar=False,
                        callback=_cb,
                    )

                    new_files = [
                        p
                        for p in destination.rglob("*")
                        if p.is_file()
                        and str(p.relative_to(destination)) not in before
                    ]

                    if not new_files:
                        raise AfcError(
                            "AFC pull finished but no new files appeared. "
                            "Unlock the iPhone and check that photos are stored on-device "
                            "(not iCloud-only placeholders)."
                        )

                    for path in new_files:
                        rel = str(path.relative_to(destination)).replace("\\", "/")
                        succeeded.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": remote,
                                "file_name": path.name,
                                "relative_path": rel,
                                "destination_path": str(path),
                                "status": "Copied",
                            }
                        )

                    if on_progress:
                        on_progress(
                            f"{remote} ({len(new_files)} files)",
                            index,
                            total,
                            "done",
                        )

                except Exception as exc:  # noqa: BLE001
                    failed.append(
                        {
                            "timestamp": _now(),
                            "device": device_name,
                            "source_folder": remote,
                            "file_name": "",
                            "relative_path": "",
                            "source_path": remote,
                            "destination_path": str(
                                destination / Path(remote).name
                            ),
                            "error": str(exc),
                            "status": "Skipped",
                        }
                    )
                    if on_progress:
                        on_progress(remote, index, total, "skipped")

    write_copy_excel_log(
        log_path,
        device_name=device_name,
        selected_folders=folder_paths,
        succeeded=succeeded,
        failed=failed,
    )
    return AfcCopyResult(
        succeeded=succeeded,
        failed=failed,
        log_path=log_path,
        selected_folders=list(folder_paths),
    )


def copy_afc_folders(
    serial: str,
    folder_paths: list[str],
    destination: Path,
    *,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> AfcCopyResult:
    if not folder_paths:
        return AfcCopyResult(selected_folders=[])
    if not afc_available():
        raise AfcError(
            "pymobiledevice3 is not installed. Run: python -m pip install pymobiledevice3"
        )
    try:
        return _run(
            _copy_folders_async(
                serial,
                folder_paths,
                destination,
                on_progress=on_progress,
            )
        )
    except AfcError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AfcError(f"AFC copy failed: {exc}") from exc
