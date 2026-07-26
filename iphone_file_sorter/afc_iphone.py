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
import time
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
FILE_TIMEOUT_FLOOR_SEC = 45.0
FILE_TIMEOUT_PER_MB_SEC = 6.0  # allow ~slow USB
FILE_TIMEOUT_CAP_SEC = 1800.0  # 30 min max per file
FILE_TIMEOUT_SMALL_SEC = 30.0  # files under 1 MB


def _format_exc(exc: BaseException) -> str:
    """Readable exception text even when str(exc) is empty."""
    text = str(exc).strip()
    name = type(exc).__name__
    if text:
        return f"{name}: {text}"
    return f"{name} (no message)"


# Real user media extensions (used when filtering PhotoData noise)
_USER_MEDIA_SUFFIXES = (
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
    ".mov",
    ".mp4",
    ".m4v",
    ".3gp",
    ".avi",
    ".mkv",
    ".aae",
    ".dng",
    ".cr2",
    ".nef",
)


def _should_skip_remote(remote_file: str) -> bool:
    """Skip iOS system caches / DB / metadata that aren't user photos/videos."""
    path = remote_file.replace("\\", "/").lower()
    markers = (
        "/caches/",
        "/cache/",
        "photodata/caches",
        "photodata/private",
        "photodata/changes/",
        "/tmp/",
        ".sqlite-wal",
        ".sqlite-shm",
        ".db-wal",
        ".db-shm",
    )
    if any(m in path for m in markers):
        return True
    if path.endswith(".sqlite") and "photodata/" in path:
        return True

    # PhotoData is mostly Photos.app databases, thumbnails, and analysis files.
    # Only keep actual media-looking files from there; DCIM holds camera originals.
    if "photodata/" in path:
        if not path.endswith(_USER_MEDIA_SUFFIXES):
            return True
    return False


def summarize_copy_rows(
    succeeded: list[dict], failed: list[dict]
) -> dict[str, int]:
    """Count copy outcomes for UI explanation."""
    summary = {
        "copied": 0,
        "already_on_disk": 0,
        "system_skipped": 0,
        "timeout": 0,
        "other_failed": 0,
    }
    for row in succeeded:
        status = str(row.get("status", ""))
        if "already" in status.lower():
            summary["already_on_disk"] += 1
        else:
            summary["copied"] += 1
    for row in failed:
        err = str(row.get("error", "")).lower()
        if "system cache" in err or "system file" in err:
            summary["system_skipped"] += 1
        elif "timed out" in err or "timeout" in err:
            summary["timeout"] += 1
        else:
            summary["other_failed"] += 1
    return summary


@dataclass
class AfcCopyResult:
    succeeded: list[dict[str, Any]] = field(default_factory=list)
    failed: list[dict[str, Any]] = field(default_factory=list)
    log_path: Path | None = None
    selected_folders: list[str] = field(default_factory=list)
    filtered_out: int = 0  # excluded by extension allow-list (not logged per-file)


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
    size = max(size_bytes, 0)
    if size < 1024 * 1024:
        return FILE_TIMEOUT_SMALL_SEC
    mb = size / (1024 * 1024)
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
            f"Could not list iPhone via Apple drivers: {_format_exc(exc)}. "
            "Unlock the iPhone, tap Trust, and ensure iTunes/Apple Devices is installed. "
            "If Access is denied when restarting the Apple service, run "
            "Restart_Apple_Mobile_Device_Service.bat as Administrator."
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
    """List all file paths under remote_dir (AFC paths). Prefer _iter_files for copy."""
    return [p async for p in _iter_files(afc, remote_dir)]


async def _iter_files(afc, remote_dir: str):
    """Yield file paths under remote_dir as they are discovered (no full pre-scan)."""
    root = (remote_dir or "").strip("/")
    try:
        if root and not await _isdir(afc, root):
            try:
                await afc.stat(root)
                yield root
            except Exception:
                return
            return
    except Exception:
        return

    # AFC walk wants "." for media root; normalize yielded paths.
    walk_root = root if root else "."
    async for dirpath, _dirnames, filenames in afc.walk(walk_root):
        norm_dir = "" if dirpath in {"", "."} else str(dirpath).lstrip("./")
        for filename in filenames:
            yield posixpath.join(norm_dir, filename) if norm_dir else filename


async def _immediate_subdirs(afc, remote: str) -> list[str]:
    """Return immediate child directory names under remote."""
    out: list[str] = []
    for name in await _listdir(afc, remote if remote else ""):
        child = _join(remote, name) if remote else name
        if await _isdir(afc, child):
            out.append(name)
    return out


async def _has_immediate_files(afc, remote: str) -> bool:
    for name in await _listdir(afc, remote if remote else ""):
        child = _join(remote, name) if remote else name
        if not await _isdir(afc, child):
            return True
    return False


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


def _ext_allowed(remote_file: str, allowed_extensions: set[str] | None) -> bool:
    if not allowed_extensions:
        return True
    return Path(remote_file).suffix.lower() in allowed_extensions


async def _copy_one_remote_file(
    afc,
    *,
    remote_file: str,
    folder_remote: str,
    destination: Path,
    device_name: str,
    logical_folder: str,
    file_index: int,
    folder_index: int,
    folder_total: int,
    on_progress: Callable[[str, int, int, str], None] | None,
    succeeded: list,
    failed: list,
    handled_remotes: set[str],
    allowed_extensions: set[str] | None = None,
    filter_stats: dict[str, int] | None = None,
) -> str:
    """
    Copy a single remote file. Returns:
      'ok' | 'skipped' | 'timeout' | 'disconnect' | 'error'
    """
    local_file = _local_path_for(destination, remote_file, folder_remote)
    label = remote_file
    # Progress: folder N/M, with running file count in the label
    prog_total = max(folder_total, 1)
    prog_index = min(folder_index, prog_total)

    if _should_skip_remote(remote_file) or not _ext_allowed(remote_file, allowed_extensions):
        handled_remotes.add(remote_file)
        # Extension / system filters are expected noise — count quietly, don't fill the log.
        if filter_stats is not None:
            filter_stats["filtered_out"] = int(filter_stats.get("filtered_out", 0)) + 1
        if on_progress and file_index % 25 == 0:
            on_progress(
                f"[folder {folder_index}/{folder_total}] scanning… skipped junk/non-matching · files {file_index}",
                prog_index,
                prog_total,
                "scanning",
            )
        return "filtered"

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
                "relative_path": str(local_file.relative_to(destination)).replace(
                    "\\", "/"
                ),
                "destination_path": str(local_file),
                "status": "Skipped (already copied)",
            }
        )
        if on_progress:
            on_progress(
                f"[folder {folder_index}/{folder_total}] {label} (on disk) · files {file_index}",
                prog_index,
                prog_total,
                "skipped",
            )
        return "skipped"

    timeout = _file_timeout_sec(remote_size)
    if on_progress:
        on_progress(
            f"[folder {folder_index}/{folder_total}] {label} · files {file_index}",
            prog_index,
            prog_total,
            "copying",
        )

    try:
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
                "relative_path": str(local_file.relative_to(destination)).replace(
                    "\\", "/"
                ),
                "destination_path": str(local_file),
                "status": "Copied",
            }
        )
        if on_progress:
            on_progress(
                f"[folder {folder_index}/{folder_total}] {label} · files {file_index}",
                prog_index,
                prog_total,
                "done",
            )
        return "ok"

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
                f"[folder {folder_index}/{folder_total}] {label} (timeout)",
                prog_index,
                prog_total,
                "skipped",
            )
        return "timeout"

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
                f"[folder {folder_index}/{folder_total}] {label} (error)",
                prog_index,
                prog_total,
                "skipped",
            )
        err = str(exc).lower()
        if any(
            token in err
            for token in ("connection", "broken", "eof", "socket", "closed")
        ):
            return "disconnect"
        return "error"


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
    allowed_extensions: set[str] | None = None,
    filter_stats: dict[str, int] | None = None,
) -> tuple[int, bool]:
    """
    Copy one media folder file-by-file with timeouts.
    Walks and copies immediately (no full-tree pre-count).
    Returns (next_file_index, needs_reconnect).
    """
    from pymobiledevice3.services.afc import AfcService

    del file_total  # progress is folder-based now
    file_index = file_index_start
    afc = AfcService(lockdown=lockdown)
    try:
        await afc.__aenter__()
    except Exception as exc:  # noqa: BLE001
        if on_progress:
            on_progress(
                f"Could not open AFC: {_format_exc(exc)}",
                0,
                1,
                "error",
            )
        return file_index, True

    try:
        # Expand broad folders (Media root / DCIM) into child dirs so the first
        # file lands quickly and progress advances per album/folder.
        if not remote:
            targets = await _immediate_subdirs(afc, "")
        else:
            subdirs = await _immediate_subdirs(afc, remote)
            only_subdirs = subdirs and not await _has_immediate_files(afc, remote)
            shallow = remote.count("/") == 0  # e.g. DCIM, Downloads
            if only_subdirs and shallow:
                targets = [_join(remote, name) for name in subdirs]
            else:
                targets = [remote]

        folder_total = max(len(targets), 1)
        if on_progress:
            on_progress(
                f"Starting copy of {logical_folder} ({folder_total} subfolder(s))",
                0,
                folder_total,
                "starting",
            )

        for folder_i, folder_remote in enumerate(targets, start=1):
            if on_progress:
                on_progress(
                    f"Opening {folder_remote}",
                    folder_i,
                    folder_total,
                    "listing",
                )

            try:
                async for remote_file in _iter_files(afc, folder_remote):
                    if remote_file in handled_remotes:
                        continue
                    file_index += 1
                    result = await _copy_one_remote_file(
                        afc,
                        remote_file=remote_file,
                        folder_remote=folder_remote,
                        destination=destination,
                        device_name=device_name,
                        logical_folder=logical_folder,
                        file_index=file_index,
                        folder_index=folder_i,
                        folder_total=folder_total,
                        on_progress=on_progress,
                        succeeded=succeeded,
                        failed=failed,
                        handled_remotes=handled_remotes,
                        allowed_extensions=allowed_extensions,
                        filter_stats=filter_stats,
                    )
                    if result in {"timeout", "disconnect"}:
                        return file_index, True
            except Exception as exc:  # noqa: BLE001
                if on_progress:
                    on_progress(
                        f"Folder error {folder_remote}: {_format_exc(exc)}",
                        folder_i,
                        folder_total,
                        "error",
                    )
                return file_index, True

        return file_index, False
    finally:
        # After a timed-out pull, AFC cleanup often throws an empty error —
        # never let that abort the whole batch.
        try:
            await afc.__aexit__(None, None, None)
        except Exception:
            pass


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
    allowed_extensions: set[str] | None = None,
    filter_stats: dict[str, int] | None = None,
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

                    if _should_skip_remote(remote_file) or not _ext_allowed(
                        remote_file, allowed_extensions
                    ):
                        if filter_stats is not None:
                            filter_stats["filtered_out"] = (
                                int(filter_stats.get("filtered_out", 0)) + 1
                            )
                        continue

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
    allowed_extensions: set[str] | None = None,
    filter_stats: dict[str, int] | None = None,
) -> AfcCopyResult:
    from pymobiledevice3.lockdown import create_using_usbmux

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    log_path = default_log_path(destination)

    succeeded: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []
    device_name = serial
    targets = _expand_copy_targets(folder_paths)
    if filter_stats is None:
        filter_stats = {"filtered_out": 0}

    # No upfront file count — that walked the whole phone and looked "stuck".
    # Copy starts immediately; progress is per folder + running file count.
    if on_progress:
        on_progress(
            f"Starting copy ({len(targets)} target(s)) — files will appear in the destination soon",
            0,
            max(len(targets), 1),
            "starting",
        )

    file_index = 0
    try:
        for target_i, logical in enumerate(targets, start=1):
            parsed = parse_logical_path(logical)
            handled_remotes: set[str] = set()
            if on_progress:
                on_progress(
                    f"Target {target_i}/{len(targets)}: {logical}",
                    target_i - 1,
                    max(len(targets), 1),
                    "starting",
                )
            # Multiple passes: after a timeout/disconnect, reconnect and continue
            # remaining files (already-handled remotes are skipped).
            completed_target = False
            for _pass in range(8):
                try:
                    async with await create_using_usbmux(serial=serial) as lockdown:
                        device_name = str(
                            getattr(lockdown, "display_name", None) or serial
                        )

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
                                file_total=0,
                                handled_remotes=handled_remotes,
                                allowed_extensions=allowed_extensions,
                                filter_stats=filter_stats,
                            )
                            if needs_reconnect:
                                continue
                            completed_target = True
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
                                file_total=max(file_index + 1, 1),
                                allowed_extensions=allowed_extensions,
                                filter_stats=filter_stats,
                            )
                            if needs_reconnect:
                                continue
                            completed_target = True
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
                        completed_target = True
                        break
                except Exception as exc:  # noqa: BLE001
                    if on_progress:
                        on_progress(
                            f"Reconnect/session issue: {_format_exc(exc)}",
                            target_i,
                            max(len(targets), 1),
                            "error",
                        )
                    await asyncio.sleep(1.0)
                    continue

            if not completed_target:
                failed.append(
                    {
                        "timestamp": _now(),
                        "device": device_name,
                        "source_folder": logical,
                        "file_name": "",
                        "relative_path": "",
                        "source_path": logical,
                        "destination_path": str(destination),
                        "error": "Gave up after repeated USB/AFC session errors",
                        "status": "Skipped",
                    }
                )
    finally:
        # Always write a log, even if a later target fails.
        try:
            write_copy_excel_log(
                log_path,
                device_name=device_name,
                selected_folders=folder_paths,
                succeeded=succeeded,
                failed=failed,
            )
        except Exception:
            log_path = None

    return AfcCopyResult(
        succeeded=succeeded,
        failed=failed,
        log_path=log_path,
        selected_folders=list(folder_paths),
        filtered_out=int(filter_stats.get("filtered_out", 0)),
    )


def copy_afc_folders(
    serial: str,
    folder_paths: list[str],
    destination: Path,
    *,
    on_progress: Callable[[str, int, int, str], None] | None = None,
    allowed_extensions: set[str] | None = None,
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
                allowed_extensions=allowed_extensions,
            )
        )
    except AfcError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AfcError(f"AFC copy failed: {_format_exc(exc)}") from exc


def _safe_dir_name(label: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in {"-", "_", " "} else "_" for ch in label)
    return cleaned.strip().replace(" ", "_")[:80] or "Content"


def _count_backup_files(backup_directory: Path) -> tuple[int, int]:
    """Return (file_count, total_bytes) under backup_directory (best-effort)."""
    files = 0
    total = 0
    try:
        for path in backup_directory.rglob("*"):
            if path.is_file():
                files += 1
                try:
                    total += path.stat().st_size
                except OSError:
                    pass
    except OSError:
        pass
    return files, total


async def _backup_selections_async(
    serial: str,
    backup_directory: Path,
    *,
    selections: list[str],
    regexes: list[str],
    password: str,
    on_progress: Callable[[str, int, int, str], None] | None,
    stall_seconds: int = 600,
) -> None:
    """
    Run a selective MobileBackup2 backup.

    Aborts if on-disk file count/size does not grow for `stall_seconds`
    (default 10 minutes) — that usually means the device session hung.
    """
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.mobilebackup2 import Mobilebackup2Service

    backup_directory.mkdir(parents=True, exist_ok=True)
    started = time.time()
    last_pct = {"v": 0}
    last_growth = {"files": -1, "bytes": -1, "t": time.time()}
    heartbeat_path = backup_directory / "_backup_heartbeat.txt"

    def _write_heartbeat(note: str) -> None:
        files, nbytes = _count_backup_files(backup_directory)
        elapsed = int(time.time() - started)
        stalled_for = int(time.time() - last_growth["t"])
        text = (
            f"{_now()}\n"
            f"note={note}\n"
            f"elapsed_sec={elapsed}\n"
            f"device_progress_pct={last_pct['v']}\n"
            f"files_on_disk={files}\n"
            f"bytes_on_disk={nbytes}\n"
            f"seconds_since_disk_growth={stalled_for}\n"
            f"selections={','.join(selections)}\n"
            f"regexes={','.join(regexes)}\n"
            "hint=If seconds_since_disk_growth keeps rising past 600 and Explorer "
            "shows no new files, stop with Ctrl+C and retry (unlock iPhone, try "
            "Contacts only, enter backup password if encrypted).\n"
        )
        try:
            heartbeat_path.write_text(text, encoding="utf-8")
        except OSError:
            pass

    def _mark_growth_if_needed() -> tuple[int, int]:
        files, nbytes = _count_backup_files(backup_directory)
        if files != last_growth["files"] or nbytes != last_growth["bytes"]:
            last_growth["files"] = files
            last_growth["bytes"] = nbytes
            last_growth["t"] = time.time()
        return files, nbytes

    def _prog(pct: float) -> None:
        fraction = max(0.0, min(float(pct), 100.0))
        # Any device progress callback counts as activity even if disk is briefly quiet.
        last_pct["v"] = int(fraction)
        last_growth["t"] = time.time()
        elapsed = int(time.time() - started)
        files, nbytes = _mark_growth_if_needed()
        mb = nbytes / (1024 * 1024)
        _write_heartbeat("device_progress")
        if on_progress:
            on_progress(
                (
                    f"iPhone backup {fraction:.0f}% · elapsed {elapsed // 60}m{elapsed % 60:02d}s · "
                    f"{files} files ({mb:.1f} MB) on disk"
                ),
                int(fraction),
                100,
                "backup",
            )

    async with await create_using_usbmux(serial=serial) as lockdown:
        async with Mobilebackup2Service(lockdown) as backup_client:
            preserve_rules = Mobilebackup2Service.resolve_backup_selection(selections)
            filter_callback = Mobilebackup2Service.combine_filter_callbacks(
                Mobilebackup2Service.selection_filter_callback(preserve_rules)
                if preserve_rules
                else None,
                Mobilebackup2Service.regex_filter_callback(regexes) if regexes else None,
            )
            if on_progress:
                on_progress(
                    (
                        "Starting selective backup. If no new files appear for 10 minutes, "
                        "the app will abort as stuck."
                    ),
                    0,
                    100,
                    "backup",
                )
            _write_heartbeat("starting")
            _mark_growth_if_needed()

            backup_task = asyncio.create_task(
                backup_client.backup(
                    full=True,
                    backup_directory=str(backup_directory),
                    progress_callback=_prog,
                    filter_callback=filter_callback,
                    password=password or "",
                ),
                name="iphone-selective-backup",
            )
            try:
                while not backup_task.done():
                    await asyncio.sleep(15.0)
                    elapsed = int(time.time() - started)
                    files, nbytes = _mark_growth_if_needed()
                    mb = nbytes / (1024 * 1024)
                    stalled_for = int(time.time() - last_growth["t"])
                    _write_heartbeat("monitor")
                    if on_progress:
                        on_progress(
                            (
                                f"monitoring · device last {last_pct['v']}% · "
                                f"elapsed {elapsed // 60}m{elapsed % 60:02d}s · "
                                f"{files} files ({mb:.1f} MB) · "
                                f"no growth for {stalled_for}s"
                            ),
                            max(last_pct["v"], 1),
                            100,
                            "backup",
                        )
                    if stalled_for >= stall_seconds:
                        backup_task.cancel()
                        try:
                            await backup_task
                        except asyncio.CancelledError:
                            pass
                        raise AfcError(
                            "Backup looks stuck: no new files and no progress for "
                            f"{stall_seconds // 60} minutes. Unlock the iPhone, keep the "
                            "screen on, unplug/replug USB, then retry with only "
                            "**Contacts** (uncheck Notes). If backups are encrypted, "
                            "enter the backup password."
                        )
                await backup_task
            finally:
                if not backup_task.done():
                    backup_task.cancel()
                    try:
                        await backup_task
                    except asyncio.CancelledError:
                        pass
                _write_heartbeat("finished_or_stopped")


async def _copy_content_categories_async(
    serial: str,
    category_ids: list[str],
    destination: Path,
    *,
    backup_password: str = "",
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> AfcCopyResult:
    from content_categories import CATEGORY_BY_ID

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    log_path = default_log_path(destination)

    succeeded: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []
    filter_stats = {"filtered_out": 0}
    device_name = serial
    selected_labels: list[str] = []

    afc_jobs: list[tuple[str, str, set[str] | None, str]] = []
    # (logical_path, dest_subdir_name, extensions, kind_label)
    backup_selections: list[str] = []
    backup_regexes: list[str] = []

    for cat_id in category_ids:
        cat = CATEGORY_BY_ID.get(cat_id)
        if cat is None:
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": cat_id,
                    "file_name": "",
                    "relative_path": "",
                    "source_path": cat_id,
                    "destination_path": str(destination),
                    "error": f"Unknown category: {cat_id}",
                    "status": "Skipped",
                }
            )
            continue
        selected_labels.append(cat.label)
        sub = _safe_dir_name(cat.label)
        if cat.kind == "afc_media":
            logical = f"{MEDIA_PREFIX}/{cat.media_remote}" if cat.media_remote else MEDIA_PREFIX
            afc_jobs.append((logical, sub, set(cat.extensions) if cat.extensions else None, cat.label))
        elif cat.kind == "afc_app":
            logical = f"{APPS_PREFIX}/{cat.bundle_id}"
            afc_jobs.append((logical, sub, set(cat.extensions) if cat.extensions else None, cat.label))
        elif cat.kind == "backup":
            backup_selections.extend(cat.backup_selections)
            backup_regexes.extend(cat.backup_regexes)

    # Deduplicate backup selections
    backup_selections = list(dict.fromkeys(backup_selections))
    backup_regexes = list(dict.fromkeys(backup_regexes))

    # Run each AFC category into its own subfolder with extension filter
    for logical, sub, exts, label in afc_jobs:
        if on_progress:
            on_progress(f"Category: {label}", 0, max(len(afc_jobs), 1), "starting")
        cat_dest = destination / sub
        cat_dest.mkdir(parents=True, exist_ok=True)
        try:
            partial = await _copy_folders_async(
                serial,
                [logical],
                cat_dest,
                on_progress=on_progress,
                allowed_extensions=exts,
                filter_stats=filter_stats,
            )
            # Re-base relative paths in log rows for clarity
            for row in partial.succeeded:
                row["source_folder"] = label
                succeeded.append(row)
            for row in partial.failed:
                row["source_folder"] = label
                failed.append(row)
            if partial.succeeded or not partial.failed:
                device_name = (
                    partial.succeeded[0].get("device", device_name)
                    if partial.succeeded
                    else device_name
                )
        except Exception as exc:  # noqa: BLE001
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": label,
                    "file_name": "",
                    "relative_path": "",
                    "source_path": logical,
                    "destination_path": str(cat_dest),
                    "error": _format_exc(exc),
                    "status": "Skipped",
                }
            )

    if backup_selections or backup_regexes:
        backup_root = destination / "iPhone_Backup_Selected"
        try:
            await _backup_selections_async(
                serial,
                backup_root,
                selections=backup_selections,
                regexes=backup_regexes,
                password=backup_password,
                on_progress=on_progress,
            )
            succeeded.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": "Selective iPhone backup",
                    "file_name": backup_root.name,
                    "relative_path": backup_root.name,
                    "destination_path": str(backup_root),
                    "status": "Copied",
                }
            )
            if on_progress:
                on_progress(
                    f"Backup saved under {backup_root}",
                    100,
                    100,
                    "done",
                )
        except Exception as exc:  # noqa: BLE001
            err = _format_exc(exc)
            hint = ""
            if "password" in err.lower() or "encrypt" in err.lower():
                hint = (
                    " — If the iPhone has encrypted backups enabled, enter the "
                    "backup password in the UI (or disable encrypted backup on the phone)."
                )
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": "Selective iPhone backup",
                    "file_name": "",
                    "relative_path": "",
                    "source_path": ",".join(backup_selections + backup_regexes),
                    "destination_path": str(backup_root),
                    "error": err + hint,
                    "status": "Skipped",
                }
            )

    try:
        write_copy_excel_log(
            log_path,
            device_name=device_name,
            selected_folders=selected_labels,
            succeeded=succeeded,
            failed=failed,
        )
    except Exception:
        log_path = None

    return AfcCopyResult(
        succeeded=succeeded,
        failed=failed,
        log_path=log_path,
        selected_folders=selected_labels,
        filtered_out=int(filter_stats.get("filtered_out", 0)),
    )


def copy_content_categories(
    serial: str,
    category_ids: list[str],
    destination: Path,
    *,
    backup_password: str = "",
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> AfcCopyResult:
    """Copy curated content categories (filtered media + optional selective backup)."""
    if not category_ids:
        return AfcCopyResult(selected_folders=[])
    if not afc_available():
        raise AfcError(
            "pymobiledevice3 is not installed. Run: python -m pip install pymobiledevice3"
        )
    try:
        return _run(
            _copy_content_categories_async(
                serial,
                category_ids,
                destination,
                backup_password=backup_password,
                on_progress=on_progress,
            )
        )
    except AfcError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AfcError(f"Content copy failed: {_format_exc(exc)}") from exc
