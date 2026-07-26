"""
Detect a connected iPhone/iPad in Windows Explorer (This PC) and copy
files using the Windows Shell — the same engine as File Explorer.

Copies file-by-file, skips failures, and returns detailed results for Excel logging.

Requires: pywin32 (Windows only)
"""

from __future__ import annotations

import subprocess
import sys
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterator

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

# Show Windows copy progress (like Explorer), auto-confirm prompts
COPY_FLAGS_VISIBLE = FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR


class WindowsShellError(RuntimeError):
    pass


@dataclass
class CopyResult:
    succeeded: list[dict[str, Any]] = field(default_factory=list)
    failed: list[dict[str, Any]] = field(default_factory=list)
    log_path: Path | None = None
    selected_folders: list[str] = field(default_factory=list)


@contextmanager
def _com_initialized() -> Iterator[None]:
    """
    Initialize COM for the current thread.

    Streamlit runs script code off the main thread, so win32com calls need
    explicit CoInitialize or they raise: CoInitialize has not been called.
    """
    try:
        import pythoncom  # type: ignore
    except ImportError as exc:
        raise WindowsShellError(
            "pywin32 is required on Windows. Run: python -m pip install pywin32"
        ) from exc

    # Apartment-threaded COM is required for Shell.Application CopyHere
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:
        pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _pump_messages() -> None:
    """Shell CopyHere needs a Win32 message pump or transfers never finish."""
    try:
        import pythoncom  # type: ignore

        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def _shell():
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise WindowsShellError(
            "pywin32 is required on Windows. Run: python -m pip install pywin32"
        ) from exc
    return win32com.client.Dispatch("Shell.Application")


def _worker_script() -> Path:
    return Path(__file__).resolve().with_name("shell_copy_worker.py")


def _copy_folder_via_worker(
    device_name: str,
    folder_rel: str,
    destination: Path,
    *,
    timeout_sec: float = 7200.0,
    on_progress: Callable[[str, int, int, str], None] | None = None,
    index: int = 1,
    total: int = 1,
) -> tuple[bool, int, str]:
    """
    Copy one folder in a separate Python process with a real message pump.
    Returns (ok, file_count, detail).
    """
    worker = _worker_script()
    if not worker.exists():
        raise WindowsShellError(f"Missing worker script: {worker}")

    cmd = [
        sys.executable,
        str(worker),
        device_name,
        folder_rel,
        str(destination),
        str(timeout_sec),
    ]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    file_count = 0
    detail = ""
    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.strip()
        if not line:
            continue
        if line.startswith("COPY_PROGRESS"):
            # COPY_PROGRESS files=12 elapsed=5
            parts = dict(
                p.split("=", 1) for p in line.split()[1:] if "=" in p
            )
            file_count = int(parts.get("files", "0"))
            elapsed = parts.get("elapsed", "?")
            if on_progress:
                on_progress(
                    f"{folder_rel} ({file_count} files, {elapsed}s)",
                    index,
                    total,
                    "copying folder",
                )
        elif line.startswith("COPY_DONE"):
            parts = dict(
                p.split("=", 1) for p in line.split()[1:] if "=" in p
            )
            if "files" in parts:
                file_count = int(parts["files"])
            detail = line
        elif line.startswith("COPY_FAIL"):
            detail = line
        elif line.startswith("COPY_START"):
            if on_progress:
                on_progress(folder_rel, index, total, "starting")

    rc = proc.wait()
    dest_folder = Path(destination) / folder_rel
    actual = _count_files(dest_folder)
    if actual > file_count:
        file_count = actual
    if rc == 0 and file_count > 0:
        return True, file_count, detail or "copied"
    if file_count > 0:
        return True, file_count, detail or "partial copy"
    return False, 0, detail or f"Worker failed with exit code {rc}"


def is_windows() -> bool:
    import sys

    return sys.platform.startswith("win")


def list_portable_apple_devices() -> list[str]:
    """Return device names under This PC that look like iPhone/iPad."""
    if not is_windows():
        return []

    try:
        with _com_initialized():
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
    except WindowsShellError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise WindowsShellError(f"Failed to list devices: {exc}") from exc


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
    try:
        with _com_initialized():
            root, _ = find_internal_storage(device_name)
            return sorted(list_child_names(root))
    except WindowsShellError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise WindowsShellError(f"Failed to list iPhone folders: {exc}") from exc


def _immediate_child_folder_nodes(folder: Any, parent_rel: str) -> list[dict[str, Any]]:
    """
    List only immediate child folders (no recursion).

    Important: do not walk into media folders by default — on iPhone MTP,
    enumerating files inside each dated folder can hang for a long time.
    """
    nodes: list[dict[str, Any]] = []
    for item in list(folder.Items()):
        if not _is_folder_item(item):
            continue
        child_name = str(item.Name)
        child_rel = f"{parent_rel}/{child_name}" if parent_rel else child_name
        nodes.append({"name": child_name, "path": child_rel, "children": []})
    nodes.sort(key=lambda n: n["name"].lower())
    return nodes


def list_iphone_folder_tree(device_name: str, *, max_depth: int = 1) -> dict[str, Any]:
    """
    Return a folder tree under Internal Storage.

    Default max_depth=1 loads only top-level folders (fast).
    Deeper levels should be loaded lazily with list_iphone_child_folders().
    """
    try:
        with _com_initialized():
            root, media_label = find_internal_storage(device_name)
            if max_depth <= 0:
                return {"name": media_label, "path": "", "children": []}

            children = _immediate_child_folder_nodes(root, "")
            # max_depth > 1 is supported but can be very slow on iPhone MTP
            if max_depth > 1:
                for child in children:
                    child_folder = get_folder_by_rel_path(root, child["path"])
                    child["children"] = _immediate_child_folder_nodes(
                        child_folder, child["path"]
                    )
            return {"name": media_label, "path": "", "children": children}
    except WindowsShellError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise WindowsShellError(f"Failed to build iPhone folder tree: {exc}") from exc


def list_iphone_child_folders(device_name: str, parent_rel_path: str) -> list[dict[str, Any]]:
    """Lazy-load immediate subfolders under one parent path."""
    try:
        with _com_initialized():
            root, _media_label = find_internal_storage(device_name)
            parent = get_folder_by_rel_path(root, parent_rel_path)
            return _immediate_child_folder_nodes(parent, parent_rel_path)
    except WindowsShellError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise WindowsShellError(
            f"Failed to list subfolders for '{parent_rel_path or 'root'}': {exc}"
        ) from exc


def get_folder_by_rel_path(media_root: Any, rel_path: str) -> Any:
    """Resolve a relative folder path under media_root. '' returns media_root."""
    if not rel_path:
        return media_root
    folder = media_root
    for part in rel_path.replace("\\", "/").split("/"):
        if part:
            folder = get_child_folder(folder, part)
    return folder


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
    media_root: Any, folder_rel: str, file_rel: str
) -> tuple[Any, Any]:
    """Return (file_item, parent_folder) for file_rel under folder_rel."""
    folder = get_folder_by_rel_path(media_root, folder_rel)
    parts = file_rel.replace("\\", "/").split("/")
    for part in parts[:-1]:
        if part:
            folder = get_child_folder(folder, part)

    filename = parts[-1]
    for item in folder.Items():
        if str(item.Name) == filename and not _is_folder_item(item):
            return item, folder
    source = f"{folder_rel}/{file_rel}" if folder_rel else file_rel
    raise WindowsShellError(f"File not found on iPhone: {source}")


def _wait_for_file(
    path: Path,
    timeout_sec: float = 45.0,
    on_tick: Callable[[float, float], None] | None = None,
) -> tuple[bool, str]:
    """
    Wait until a copied file exists with a stable non-zero size.
    Returns (ok, reason_if_not_ok).
    """
    start = time.time()
    last_size = -1
    stable_rounds = 0
    saw_exists = False

    while time.time() - start < timeout_sec:
        _pump_messages()
        elapsed = time.time() - start
        if on_tick is not None:
            on_tick(elapsed, timeout_sec)

        if path.exists() and path.is_file():
            saw_exists = True
            try:
                size = path.stat().st_size
            except OSError:
                time.sleep(0.2)
                continue

            # Accept any non-empty file that stayed the same size briefly
            if size > 0 and size == last_size:
                stable_rounds += 1
                if stable_rounds >= 2:
                    return True, ""
            else:
                stable_rounds = 0
                last_size = size
        time.sleep(0.2)

    if saw_exists:
        try:
            size = path.stat().st_size if path.exists() else 0
        except OSError:
            size = 0
        if size > 0:
            return True, ""
        return False, "File appeared but stayed empty (likely MTP/copy failure)"
    return False, "Timed out — file did not appear at destination"


def _count_files(folder: Path) -> int:
    if not folder.exists():
        return 0
    return sum(1 for p in folder.rglob("*") if p.is_file())


def _wait_for_folder_copy(
    dest_folder: Path,
    *,
    timeout_sec: float = 7200.0,
    on_tick: Callable[[int, float], None] | None = None,
) -> tuple[bool, int, str]:
    """
    Wait until a folder copy finishes (file count stops changing).
    Returns (ok, file_count, reason).
    """
    start = time.time()
    last_count = -1
    stable_rounds = 0

    while time.time() - start < timeout_sec:
        _pump_messages()
        count = _count_files(dest_folder)
        elapsed = time.time() - start
        if on_tick is not None:
            on_tick(count, elapsed)

        if count > 0 and count == last_count:
            stable_rounds += 1
            # ~5 seconds with no new files
            if stable_rounds >= 5:
                return True, count, ""
        else:
            stable_rounds = 0
            last_count = count
        time.sleep(1.0)

    count = _count_files(dest_folder)
    if count > 0:
        return True, count, "Timed out but some files arrived"
    return False, 0, "Timed out — folder copy produced no files"


def _resolve_folder_item(media_root: Any, folder_rel: str) -> Any:
    """Return the Shell folder item for a relative path under media_root."""
    if not folder_rel:
        raise WindowsShellError("Cannot resolve empty folder path as a single item.")
    parts = [p for p in folder_rel.replace("\\", "/").split("/") if p]
    parent = media_root
    for part in parts[:-1]:
        parent = get_child_folder(parent, part)
    name = parts[-1]
    for item in parent.Items():
        if str(item.Name) == name and _is_folder_item(item):
            return item
    raise WindowsShellError(f"Folder not found on iPhone: {folder_rel}")


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def copy_folders_bulk(
    device_name: str,
    folder_names: list[str],
    destination: Path,
    *,
    log_path: Path | None = None,
    folder_timeout_sec: float = 7200.0,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> CopyResult:
    """
    Copy each selected folder as a whole using the Windows copy dialog.

    Much faster/more reliable than file-by-file MTP copies for large folders.
    If Windows skips bad files, remaining files still copy.
    """
    if len(folder_names) == 0:
        return CopyResult(selected_folders=[])

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    if log_path is None:
        log_path = default_log_path(destination)

    succeeded: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []

    # Resolve target folder names (may need a short COM call)
    with _com_initialized():
        media_root, media_label = find_internal_storage(device_name)
        targets: list[str] = []
        for folder_rel in folder_names:
            if folder_rel == "":
                targets.extend(list_child_names(media_root))
            else:
                targets.append(folder_rel)

    seen: set[str] = set()
    unique_targets: list[str] = []
    for t in targets:
        if t not in seen:
            seen.add(t)
            unique_targets.append(t)

    total = len(unique_targets)
    for index, folder_rel in enumerate(unique_targets, start=1):
        dest_folder = destination / folder_rel
        try:
            # Separate process with message pump — required for MTP CopyHere
            ok, count, reason = _copy_folder_via_worker(
                device_name,
                folder_rel,
                destination,
                timeout_sec=folder_timeout_sec,
                on_progress=on_progress,
                index=index,
                total=total,
            )
            if not ok:
                raise WindowsShellError(reason or "Folder copy failed")

            if dest_folder.exists():
                for path in dest_folder.rglob("*"):
                    if path.is_file():
                        rel = str(path.relative_to(dest_folder)).replace("\\", "/")
                        succeeded.append(
                            {
                                "timestamp": _now(),
                                "device": device_name,
                                "source_folder": folder_rel,
                                "file_name": path.name,
                                "relative_path": rel,
                                "destination_path": str(path),
                                "status": "Copied",
                            }
                        )

            if on_progress:
                on_progress(
                    f"{folder_rel} ({count} files)",
                    index,
                    total,
                    "done",
                )

        except Exception as exc:  # noqa: BLE001
            failed.append(
                {
                    "timestamp": _now(),
                    "device": device_name,
                    "source_folder": folder_rel,
                    "file_name": "",
                    "relative_path": "",
                    "source_path": f"{device_name}/{media_label}/{folder_rel}",
                    "destination_path": str(dest_folder),
                    "error": str(exc),
                    "status": "Skipped",
                }
            )
            if on_progress:
                on_progress(folder_rel, index, total, "skipped")
            continue

    write_copy_excel_log(
        log_path,
        device_name=device_name,
        selected_folders=unique_targets,
        succeeded=succeeded,
        failed=failed,
    )
    return CopyResult(
        succeeded=succeeded,
        failed=failed,
        log_path=Path(log_path),
        selected_folders=list(unique_targets),
    )


def copy_folders_file_by_file(
    device_name: str,
    folder_names: list[str],
    destination: Path,
    *,
    log_path: Path | None = None,
    file_timeout_sec: float = 45.0,
    on_progress: Callable[[str, int, int, str], None] | None = None,
) -> CopyResult:
    """
    Copy all files under the selected iPhone folders to destination.

    folder_names are relative paths under Internal Storage.
    Use "" to copy the entire media root.

    - Works file-by-file
    - On error/timeout for one file: log it, skip, continue
    - Writes an Excel log with failed/successful details
    """
    if len(folder_names) == 0:
        return CopyResult(selected_folders=[])

    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    if log_path is None:
        log_path = default_log_path(destination)

    with _com_initialized():
        shell = _shell()
        media_root, media_label = find_internal_storage(device_name)

        # Build job list: (folder_rel, file_rel_within_folder)
        normalized_jobs: list[tuple[str, str]] = []
        failed: list[dict[str, Any]] = []
        for folder_rel in folder_names:
            display_folder = folder_rel or media_label
            try:
                folder = get_folder_by_rel_path(media_root, folder_rel)
            except WindowsShellError as exc:
                failed.append(
                    {
                        "timestamp": _now(),
                        "device": device_name,
                        "source_folder": display_folder,
                        "file_name": "",
                        "relative_path": "",
                        "source_path": (
                            f"{device_name}/{media_label}/{folder_rel}"
                            if folder_rel
                            else f"{device_name}/{media_label}"
                        ),
                        "destination_path": str(
                            destination / folder_rel if folder_rel else destination
                        ),
                        "error": str(exc),
                        "status": "Skipped",
                    }
                )
                continue

            for rel in _enumerate_relative_files(folder):
                normalized_jobs.append((folder_rel, rel))

        succeeded: list[dict[str, Any]] = []
        total = len(normalized_jobs)

        for index, (folder_rel, rel) in enumerate(normalized_jobs, start=1):
            file_name = Path(rel).name
            display_folder = folder_rel or media_label
            if folder_rel:
                source_path = f"{device_name}/{media_label}/{folder_rel}/{rel}"
                dest_file = destination / folder_rel / Path(rel)
                label = f"{folder_rel}/{rel}"
            else:
                source_path = f"{device_name}/{media_label}/{rel}"
                dest_file = destination / Path(rel)
                label = rel

            dest_file.parent.mkdir(parents=True, exist_ok=True)

            if on_progress:
                on_progress(label, index, total, "copying")

            try:
                item, _parent = _resolve_file_item(media_root, folder_rel, rel)
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

                # Visible Windows dialog helps MTP transfers; user can Skip bad files
                dest_ns.CopyHere(item, COPY_FLAGS_VISIBLE)

                def _tick(elapsed: float, timeout: float, _label=label, _i=index, _t=total):
                    if on_progress:
                        on_progress(
                            f"{_label} (waiting {int(elapsed)}s / {int(timeout)}s)",
                            _i,
                            _t,
                            "waiting",
                        )

                ok, reason = _wait_for_file(
                    dest_file,
                    timeout_sec=file_timeout_sec,
                    on_tick=_tick,
                )
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
                        "source_folder": display_folder,
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
                        "source_folder": display_folder,
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
        selected_folders=[f or media_label for f in folder_names],
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
