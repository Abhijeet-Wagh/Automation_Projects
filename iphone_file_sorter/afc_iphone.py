"""
Copy real files from iPhone using Apple AFC (via pymobiledevice3).

This bypasses Windows MTP, which often creates empty folders and never
transfers file contents.

AFC media root = /var/mobile/Media (DCIM, Downloads, Recordings, …).
App sandboxes (WhatsApp, etc.) are accessed separately via house_arrest when
iOS allows it — many apps block that, which is why Explorer MTP folders are
not always visible here.
"""

from __future__ import annotations

import asyncio
import posixpath
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any, Callable

from copy_log import default_log_path, write_copy_excel_log


class AfcError(RuntimeError):
    pass


# Logical tree prefixes (not raw AFC paths)
MEDIA_PREFIX = "media"
APPS_PREFIX = "apps"
# Empty path = virtual tree root (matches folder_tree.parent_path behavior)
ROOT_PATH = ""

# Apps we try to expose via house_arrest (Documents / container).
# Many apps deny this; failures are skipped quietly.
KNOWN_APP_BUNDLES: list[tuple[str, str]] = [
    ("WhatsApp", "net.whatsapp.WhatsApp"),
    ("WhatsApp Business", "net.whatsapp.WhatsAppSMB"),
    ("Telegram", "ph.telegra.Telegraph"),
]

# Per-file timeout floor / scale (seconds). Large videos need more time.
FILE_TIMEOUT_FLOOR_SEC = 90.0
FILE_TIMEOUT_PER_MB_SEC = 6.0  # allow ~slow USB
FILE_TIMEOUT_CAP_SEC = 1800.0  # 30 min max per file


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


def _file_timeout_sec(size_bytes: int) -> float:
    mb = max(size_bytes, 0) / (1024 * 1024)
    return min(
        FILE_TIMEOUT_CAP_SEC,
        max(FILE_TIMEOUT_FLOOR_SEC, mb * FILE_TIMEOUT_PER_MB_SEC),
    )


def parse_logical_path(logical: str) -> dict[str, str]:
    """
    Parse UI tree path into a copy target.

    Returns dict with keys:
      kind: 'media' | 'app' | 'virtual'
      remote: AFC path relative to service root
      bundle_id: for app kind
    """
    path = (logical or "").strip().strip("/")
    if path in {"", ROOT_PATH}:
        return {"kind": "virtual", "remote": "", "bundle_id": ""}
    if path == MEDIA_PREFIX:
        return {"kind": "media", "remote": "", "bundle_id": ""}
    if path == APPS_PREFIX:
        return {"kind": "virtual", "remote": "", "bundle_id": ""}

    if path.startswith(MEDIA_PREFIX + "/"):
        return {
            "kind": "media",
            "remote": path[len(MEDIA_PREFIX) + 1 :],
            "bundle_id": "",
        }

    if path.startswith(APPS_PREFIX + "/"):
        rest = path[len(APPS_PREFIX) + 1 :]
        if "/" in rest:
            bundle_id, remote = rest.split("/", 1)
        else:
            bundle_id, remote = rest, ""
        return {"kind": "app", "remote": remote, "bundle_id": bundle_id}

    # Backward-compatible raw AFC paths from older UI sessions (e.g. DCIM/127APPLE)
    return {"kind": "media", "remote": path, "bundle_id": ""}


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


async def _stat_size(afc, path: str) -> int:
    try:
        info = await afc.stat(path if path else ".")
        return int(info.get("st_size") or 0)
    except Exception:
        return 0


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


async def _build_media_subtree(afc, remote: str, logical: str, name: str, depth: int, max_depth: int) -> dict:
    node: dict[str, Any] = {
        "name": name,
        "path": logical,
        "children": [],
        "source": "media",
    }
    if depth >= max_depth:
        return node
    for child_name in await _listdir(afc, remote):
        child_remote = _join(remote, child_name)
        if await _isdir(afc, child_remote):
            child_logical = _join(logical, child_name)
            node["children"].append(
                await _build_media_subtree(
                    afc, child_remote, child_logical, child_name, depth + 1, max_depth
                )
            )
    return node


async def _build_app_subtree(afc, remote: str, logical: str, name: str, depth: int, max_depth: int, bundle_id: str) -> dict:
    node: dict[str, Any] = {
        "name": name,
        "path": logical,
        "children": [],
        "source": "app",
        "bundle_id": bundle_id,
    }
    if depth >= max_depth:
        return node
    for child_name in await _listdir(afc, remote if remote else "."):
        child_remote = _join(remote, child_name) if remote else child_name
        if await _isdir(afc, child_remote):
            child_logical = _join(logical, child_name)
            node["children"].append(
                await _build_app_subtree(
                    afc,
                    child_remote,
                    child_logical,
                    child_name,
                    depth + 1,
                    max_depth,
                    bundle_id,
                )
            )
    return node


async def _try_app_tree(lockdown, label: str, bundle_id: str, max_depth: int) -> dict | None:
    from pymobiledevice3.services.house_arrest import HouseArrestService

    for documents_only in (False, True):
        try:
            async with await HouseArrestService.create(
                lockdown=lockdown,
                bundle_id=bundle_id,
                documents_only=documents_only,
            ) as ha:
                logical = f"{APPS_PREFIX}/{bundle_id}"
                suffix = " (Documents)" if documents_only else ""
                return await _build_app_subtree(
                    ha,
                    "",
                    logical,
                    f"{label}{suffix}",
                    0,
                    max_depth,
                    bundle_id,
                )
        except Exception:
            continue
    return None


async def _build_tree_async(serial: str, max_depth: int = 2) -> dict:
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.afc import AfcService

    async with await create_using_usbmux(serial=serial) as lockdown:
        media_children: list[dict] = []
        async with AfcService(lockdown=lockdown) as afc:
            for child_name in await _listdir(afc, ""):
                child_remote = child_name
                if await _isdir(afc, child_remote):
                    media_children.append(
                        await _build_media_subtree(
                            afc,
                            child_remote,
                            f"{MEDIA_PREFIX}/{child_name}",
                            child_name,
                            1,
                            max_depth,
                        )
                    )

        media_node: dict[str, Any] = {
            "name": "Media (DCIM, Downloads, Recordings, …)",
            "path": MEDIA_PREFIX,
            "children": media_children,
            "source": "media",
        }

        app_children: list[dict] = []
        for label, bundle_id in KNOWN_APP_BUNDLES:
            node = await _try_app_tree(lockdown, label, bundle_id, max_depth)
            if node is not None:
                app_children.append(node)

        children: list[dict] = [media_node]
        if app_children:
            children.append(
                {
                    "name": "Apps (WhatsApp / Telegram if iOS allows)",
                    "path": APPS_PREFIX,
                    "children": app_children,
                    "source": "apps",
                }
            )
        else:
            # Keep an empty Apps node so the UI can explain the limitation
            children.append(
                {
                    "name": "Apps (none accessible via USB — see note below)",
                    "path": APPS_PREFIX,
                    "children": [],
                    "source": "apps",
                }
            )

        return {
            "name": "iPhone",
            "path": "",
            "children": children,
            "source": "root",
        }


def list_afc_folder_tree(serial: str, *, max_depth: int = 2) -> dict:
    try:
        return _run(_build_tree_async(serial, max_depth=max_depth))
    except AfcError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AfcError(f"Failed to read iPhone folders via AFC: {exc}") from exc


async def _collect_files(afc, remote_dir: str) -> list[str]:
    """List all file paths under remote_dir (AFC paths)."""
    files: list[str] = []
    root = remote_dir if remote_dir else "."
    try:
        if not await _isdir(afc, root if root != "." else ""):
            # Single file?
            try:
                await afc.stat(remote_dir)
                return [remote_dir]
            except Exception:
                return []
    except Exception:
        return []

    async for dirpath, _dirnames, filenames in afc.walk(root if root != "." else ""):
        for filename in filenames:
            files.append(posixpath.join(dirpath, filename) if dirpath not in {"", "."} else filename)
    return files


def _local_path_for(destination: Path, remote_file: str, folder_remote: str) -> Path:
    """
    Map a remote file under folder_remote into destination preserving relative layout.
    Example: folder DCIM/127APPLE, file DCIM/127APPLE/IMG.JPG
      -> destination/127APPLE/IMG.JPG  (basename of folder + relative)
    """
    remote_file = remote_file.replace("\\", "/").lstrip("/")
    folder_remote = (folder_remote or "").replace("\\", "/").strip("/")

    if folder_remote and (remote_file == folder_remote or remote_file.startswith(folder_remote + "/")):
        rel = remote_file[len(folder_remote) :].lstrip("/")
        folder_name = PurePosixPath(folder_remote).name or "media"
        return destination / folder_name / rel if rel else destination / folder_name

    # Fallback: keep full remote-relative path
    return destination / remote_file


async def _pull_file(afc, remote_file: str, local_file: Path) -> None:
    local_file.parent.mkdir(parents=True, exist_ok=True)
    # Pull into parent directory (AFC places basename there)
    await afc.pull(
        remote_file,
        str(local_file.parent),
        ignore_errors=False,
        progress_bar=False,
    )
    # If pull named the file correctly we're done; if dest was a file path mismatch, rename
    pulled = local_file.parent / Path(remote_file).name
    if pulled != local_file and pulled.exists():
        if local_file.exists():
            local_file.unlink()
        pulled.replace(local_file)


async def _copy_media_folder(
    lockdown,
    remote: str,
    destination: Path,
    *,
    device_name: str,
    logical_folder: str,
    on_progress: Callable[[str, int, int, str], None] | None,
    succeeded: list,
    failed: list,
    file_index_start: int,
    file_total: int,
    handled_remotes: set[str],
) -> tuple[int, bool]:
    """
    Copy one media folder file-by-file with timeouts.
    Returns (next_file_index, needs_reconnect).
    """
    from pymobiledevice3.services.afc import AfcService

    file_index = file_index_start
    async with AfcService(lockdown=lockdown) as afc:
        # If remote is empty, copy each top-level media directory
        targets: list[str]
        if not remote:
            targets = [
                name
                for name in await _listdir(afc, "")
                if await _isdir(afc, name)
            ]
        else:
            targets = [remote]

        for folder_remote in targets:
            files = await _collect_files(afc, folder_remote)
            # Adjust total if we discovered more precisely for this folder only
            # (file_total may be 0 when unknown — then use local count)
            local_total = file_total if file_total > 0 else max(len(files), 1)

            for remote_file in files:
                if remote_file in handled_remotes:
                    continue
                file_index += 1
                local_file = _local_path_for(destination, remote_file, folder_remote)
                label = remote_file

                # Resume: skip existing same-size files
                try:
                    remote_size = await _stat_size(afc, remote_file)
                except Exception:
                    remote_size = 0

                if local_file.is_file() and remote_size > 0 and local_file.stat().st_size == remote_size:
                    handled_remotes.add(remote_file)
                    succeeded.append(
                        {
                            "timestamp": _now(),
                            "device": device_name,
                            "source_folder": logical_folder,
                            "file_name": local_file.name,
                            "relative_path": str(local_file.relative_to(destination)).replace("\\", "/"),
                            "destination_path": str(local_file),
                            "status": "Skipped (already copied)",
                        }
                    )
                    if on_progress:
                        on_progress(
                            f"{label} (already on disk)",
                            file_index,
                            local_total,
                            "skipped",
                        )
                    continue

                timeout = _file_timeout_sec(remote_size)
                if on_progress:
                    on_progress(label, file_index, local_total, "copying")

                try:
                    # Remove partial leftovers from a previous interrupted pull
                    if local_file.exists():
                        try:
                            local_file.unlink()
                        except OSError:
                            pass
                    await asyncio.wait_for(
                        _pull_file(afc, remote_file, local_file),
                        timeout=timeout,
                    )
                    if not local_file.is_file() or local_file.stat().st_size == 0:
                        # Some pulls write basename only — check alternate
                        alt = local_file.parent / Path(remote_file).name
                        if alt.is_file() and alt != local_file:
                            local_file = alt
                    if not local_file.is_file():
                        raise AfcError("File missing after pull")

                    handled_remotes.add(remote_file)
                    succeeded.append(
                        {
                            "timestamp": _now(),
                            "device": device_name,
                            "source_folder": logical_folder,
                            "file_name": local_file.name,
                            "relative_path": str(
                                local_file.relative_to(destination)
                            ).replace("\\", "/"),
                            "destination_path": str(local_file),
                            "status": "Copied",
                        }
                    )
                    if on_progress:
                        on_progress(label, file_index, local_total, "done")

                except asyncio.TimeoutError:
                    handled_remotes.add(remote_file)
                    failed.append(
                        {
                            "timestamp": _now(),
                            "device": device_name,
                            "source_folder": logical_folder,
                            "file_name": Path(remote_file).name,
                            "relative_path": remote_file,
                            "source_path": remote_file,
                            "destination_path": str(local_file),
                            "error": f"Timed out after {int(timeout)}s — skipped",
                            "status": "Skipped",
                        }
                    )
                    if on_progress:
                        on_progress(
                            f"{label} (timeout — skipped)",
                            file_index,
                            local_total,
                            "skipped",
                        )
                    # Connection may be unhealthy after cancel — reconnect
                    return file_index, True

                except Exception as exc:  # noqa: BLE001
                    handled_remotes.add(remote_file)
                    failed.append(
                        {
                            "timestamp": _now(),
                            "device": device_name,
                            "source_folder": logical_folder,
                            "file_name": Path(remote_file).name,
                            "relative_path": remote_file,
                            "source_path": remote_file,
                            "destination_path": str(local_file),
                            "error": str(exc),
                            "status": "Skipped",
                        }
                    )
                    if on_progress:
                        on_progress(
                            f"{label} (error — skipped)",
                            file_index,
                            local_total,
                            "skipped",
                        )
                    # Soft errors: keep going; hard disconnect: reconnect
                    err = str(exc).lower()
                    if any(
                        token in err
                        for token in ("connection", "broken", "eof", "socket", "closed")
                    ):
                        return file_index, True

    return file_index, False


async def _copy_app_folder(
    lockdown,
    bundle_id: str,
    remote: str,
    destination: Path,
    *,
    device_name: str,
    logical_folder: str,
    on_progress: Callable[[str, int, int, str], None] | None,
    succeeded: list,
    failed: list,
    file_index_start: int,
    file_total: int,
) -> tuple[int, bool]:
    from pymobiledevice3.services.house_arrest import HouseArrestService

    file_index = file_index_start
    last_error: Exception | None = None

    for documents_only in (False, True):
        try:
            async with await HouseArrestService.create(
                lockdown=lockdown,
                bundle_id=bundle_id,
                documents_only=documents_only,
            ) as ha:
                folder_remote = remote
                files = await _collect_files(ha, folder_remote if folder_remote else ".")
                local_total = file_total if file_total > 0 else max(len(files), 1)
                app_dest_root = destination / f"App_{bundle_id}"

                for remote_file in files:
                    file_index += 1
                    # Keep path under App_<bundle>/...
                    rel = remote_file.lstrip("./")
                    local_file = app_dest_root / rel
                    label = f"{bundle_id}:{remote_file}"

                    try:
                        remote_size = await _stat_size(ha, remote_file)
                    except Exception:
                        remote_size = 0

                    if local_file.is_file() and remote_size > 0 and local_file.stat().st_size == remote_size:
                        succeeded.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": logical_folder,
                                "file_name": local_file.name,
                                "relative_path": str(
                                    local_file.relative_to(destination)
                                ).replace("\\", "/"),
                                "destination_path": str(local_file),
                                "status": "Skipped (already copied)",
                            }
                        )
                        if on_progress:
                            on_progress(label, file_index, local_total, "skipped")
                        continue

                    timeout = _file_timeout_sec(remote_size)
                    if on_progress:
                        on_progress(label, file_index, local_total, "copying")
                    try:
                        await asyncio.wait_for(
                            _pull_file(ha, remote_file, local_file),
                            timeout=timeout,
                        )
                        if not local_file.is_file():
                            alt = local_file.parent / Path(remote_file).name
                            if alt.is_file():
                                local_file = alt
                        if not local_file.is_file():
                            raise AfcError("File missing after pull")
                        succeeded.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": logical_folder,
                                "file_name": local_file.name,
                                "relative_path": str(
                                    local_file.relative_to(destination)
                                ).replace("\\", "/"),
                                "destination_path": str(local_file),
                                "status": "Copied",
                            }
                        )
                        if on_progress:
                            on_progress(label, file_index, local_total, "done")
                    except asyncio.TimeoutError:
                        failed.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": logical_folder,
                                "file_name": Path(remote_file).name,
                                "relative_path": remote_file,
                                "source_path": remote_file,
                                "destination_path": str(local_file),
                                "error": f"Timed out after {int(timeout)}s — skipped",
                                "status": "Skipped",
                            }
                        )
                        if on_progress:
                            on_progress(label, file_index, local_total, "skipped")
                        return file_index, True
                    except Exception as exc:  # noqa: BLE001
                        failed.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": logical_folder,
                                "file_name": Path(remote_file).name,
                                "relative_path": remote_file,
                                "source_path": remote_file,
                                "destination_path": str(local_file),
                                "error": str(exc),
                                "status": "Skipped",
                            }
                        )
                        if on_progress:
                            on_progress(label, file_index, local_total, "skipped")
                return file_index, False
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            continue

    failed.append(
        {
            "timestamp": _now(),
            "device": device_name,
            "source_folder": logical_folder,
            "file_name": "",
            "relative_path": "",
            "source_path": logical_folder,
            "destination_path": str(destination),
            "error": f"Could not open app container {bundle_id}: {last_error}",
            "status": "Skipped",
        }
    )
    return file_index, False


async def _estimate_file_total(lockdown, folder_paths: list[str]) -> int:
    """Best-effort count of files for progress (media folders only)."""
    from pymobiledevice3.services.afc import AfcService

    total = 0
    try:
        async with AfcService(lockdown=lockdown) as afc:
            for logical in folder_paths:
                parsed = parse_logical_path(logical)
                if parsed["kind"] != "media":
                    continue
                remote = parsed["remote"]
                if not remote:
                    for name in await _listdir(afc, ""):
                        if await _isdir(afc, name):
                            total += len(await _collect_files(afc, name))
                else:
                    total += len(await _collect_files(afc, remote))
    except Exception:
        return 0
    return total


def _expand_copy_targets(folder_paths: list[str]) -> list[str]:
    """Expand virtual tree nodes into concrete media/app targets."""
    expanded: list[str] = []
    for logical in folder_paths:
        key = (logical or "").strip("/")
        if key in {"", "iphone"}:
            expanded.append(MEDIA_PREFIX)
            for _label, bundle_id in KNOWN_APP_BUNDLES:
                expanded.append(f"{APPS_PREFIX}/{bundle_id}")
        elif key == APPS_PREFIX:
            for _label, bundle_id in KNOWN_APP_BUNDLES:
                expanded.append(f"{APPS_PREFIX}/{bundle_id}")
        else:
            expanded.append(logical)

    seen: set[str] = set()
    targets: list[str] = []
    for item in expanded:
        if item not in seen:
            seen.add(item)
            targets.append(item)
    return targets


async def _copy_folders_async(
    serial: str,
    folder_paths: list[str],
    destination: Path,
    *,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> AfcCopyResult:
    from pymobiledevice3.lockdown import create_using_usbmux

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    log_path = default_log_path(destination)

    succeeded: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []
    device_name = serial
    targets = _expand_copy_targets(folder_paths)

    # Fresh lockdown per folder keeps USB sessions healthy after timeouts.
    async with await create_using_usbmux(serial=serial) as lockdown:
        device_name = str(getattr(lockdown, "display_name", None) or serial)
        if on_progress:
            on_progress("Counting files…", 0, 1, "preparing")
        file_total = await _estimate_file_total(lockdown, targets)
    if file_total <= 0:
        file_total = 1

    file_index = 0
    for logical in targets:
        parsed = parse_logical_path(logical)
        handled_remotes: set[str] = set()
        # Multiple passes: after a timeout/disconnect, reconnect and continue
        # remaining files (already-handled remotes are skipped).
        for _pass in range(8):
            async with await create_using_usbmux(serial=serial) as lockdown:
                device_name = str(getattr(lockdown, "display_name", None) or serial)

                if parsed["kind"] == "media":
                    file_index, needs_reconnect = await _copy_media_folder(
                        lockdown,
                        parsed["remote"],
                        destination,
                        device_name=device_name,
                        logical_folder=logical,
                        on_progress=on_progress,
                        succeeded=succeeded,
                        failed=failed,
                        file_index_start=file_index,
                        file_total=file_total,
                        handled_remotes=handled_remotes,
                    )
                    if needs_reconnect:
                        continue
                    break

                if parsed["kind"] == "app" and parsed["bundle_id"]:
                    file_index, needs_reconnect = await _copy_app_folder(
                        lockdown,
                        parsed["bundle_id"],
                        parsed["remote"],
                        destination,
                        device_name=device_name,
                        logical_folder=logical,
                        on_progress=on_progress,
                        succeeded=succeeded,
                        failed=failed,
                        file_index_start=file_index,
                        file_total=max(file_total, file_index + 1),
                    )
                    if needs_reconnect:
                        continue
                    break

                failed.append(
                    {
                        "timestamp": _now(),
                        "device": device_name,
                        "source_folder": logical,
                        "file_name": "",
                        "relative_path": "",
                        "source_path": logical,
                        "destination_path": str(destination),
                        "error": "Nothing to copy for this tree node",
                        "status": "Skipped",
                    }
                )
                break

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
