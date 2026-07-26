#!/usr/bin/env python3
"""
iPhone File Copier & Sorter — Streamlit UI

Launch:
    streamlit run app.py
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import streamlit as st

from folder_picker import pick_folder
from folder_tree import all_paths, apply_check, count_folders, selection_roots
from sort_files import CATEGORY_EXTENSIONS, categorize, iter_source_files, sort_files

_AFC_IMPORT_ERROR = ""
try:
    from afc_iphone import (
        AfcError,
        afc_available,
        copy_afc_folders,
        list_afc_devices,
        list_afc_folder_tree,
    )
except Exception as exc:  # pragma: no cover
    _AFC_IMPORT_ERROR = f"{type(exc).__name__}: {exc}"
    AfcError = RuntimeError  # type: ignore

    def afc_available() -> bool:
        return False

    def list_afc_devices():
        return []

    def list_afc_folder_tree(_serial: str, *, max_depth: int = 2):
        return {"name": "DCIM", "path": "DCIM", "children": []}

    def copy_afc_folders(*_args, **_kwargs):
        raise RuntimeError("AFC copy unavailable.")


try:
    from windows_iphone import is_windows
except Exception:  # pragma: no cover

    def is_windows() -> bool:
        return sys.platform.startswith("win")


def _python_exe() -> str:
    return sys.executable or "python"


def _pymobiledevice3_status() -> tuple[bool, str]:
    """Return (ok, detail) for the Python running this Streamlit app."""
    if _AFC_IMPORT_ERROR:
        return False, f"Could not import afc_iphone: {_AFC_IMPORT_ERROR}"
    try:
        import pymobiledevice3  # noqa: F401

        return True, f"pymobiledevice3 is available in `{_python_exe()}`"
    except ImportError as exc:
        return False, f"pymobiledevice3 not found in `{_python_exe()}`: {exc}"


def _install_pymobiledevice3() -> tuple[bool, str]:
    """
    Install pymobiledevice3 into the current interpreter.

    On Windows + Python 3.13, dependency lzfse often has no wheel and fails to
    compile. Prefer binary wheels; otherwise return guidance for a 3.12 env.
    """
    py_version = sys.version_info
    logs: list[str] = [f"Python: {_python_exe()}", f"Version: {sys.version}"]

    if py_version >= (3, 13):
        guidance = (
            "Your Anaconda Python is 3.13+. The iPhone copy library needs "
            "dependency 'lzfse', which has no Windows wheel for 3.13 and fails "
            "to compile.\n\n"
            "Create a Python 3.12 environment (recommended):\n"
            "  conda create -n iphone_copy python=3.12 -y\n"
            "  conda activate iphone_copy\n"
            "  cd %USERPROFILE%\\Documents\\Automation_Projects\\iphone_file_sorter\n"
            "  python -m pip install -r requirements.txt\n"
            "  python -m streamlit run app.py\n"
        )
        return False, guidance

    commands = [
        [_python_exe(), "-m", "pip", "install", "-U", "pip", "setuptools", "wheel"],
        # Prefer binary wheel for lzfse (avoids MSVC build)
        [
            _python_exe(),
            "-m",
            "pip",
            "install",
            "-U",
            "lzfse",
            "--only-binary=:all:",
        ],
        [_python_exe(), "-m", "pip", "install", "-U", "pymobiledevice3"],
    ]

    for cmd in commands:
        logs.append("\n$ " + " ".join(cmd))
        try:
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception as exc:  # noqa: BLE001
            logs.append(f"Failed to run command: {exc}")
            return False, "\n".join(logs)
        out = ((completed.stdout or "") + "\n" + (completed.stderr or "")).strip()
        logs.append(out or f"(exit {completed.returncode})")
        # lzfse binary install may fail on some platforms; continue and let
        # pymobiledevice3 try (AFC itself does not require lzfse).
        if completed.returncode != 0 and "pymobiledevice3" in cmd[-1]:
            logs.append(
                "\nIf this failed because of lzfse/build tools, use Python 3.12:\n"
                "  conda create -n iphone_copy python=3.12 -y\n"
                "  conda activate iphone_copy\n"
                "  python -m pip install -r requirements.txt\n"
                "  python -m streamlit run app.py"
            )
            return False, "\n".join(logs)

    return True, "\n".join(logs)


st.set_page_config(
    page_title="iPhone File Copier & Sorter",
    page_icon="📁",
    layout="centered",
)

CATEGORY_ORDER = ("Images", "Videos", "Documents", "Excel", "PDF", "Other")


def _init_state() -> None:
    defaults = {
        "copy_destination": "",
        "copy_sort_destination": "",
        "sort_source": "",
        "sort_destination": "",
        "selected_device": "",
        "afc_serial": "",
        "iphone_tree": None,
        "tree_selected": [],
        "completed_folders": [],
        "batch_size": 5,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def path_with_browse(label: str, state_key: str, placeholder: str, help_text: str) -> str:
    """
    Text path + Browse button.

    Browse cannot write directly to the text_input key after that widget is
    created, so we stash the chosen path in a pending key and apply it on the
    next run before the text_input is instantiated.
    """
    pending_key = f"_pending_path_{state_key}"
    if pending_key in st.session_state:
        st.session_state[state_key] = st.session_state.pop(pending_key)

    col_path, col_btn = st.columns([4, 1])
    with col_path:
        value = st.text_input(
            label,
            key=state_key,
            placeholder=placeholder,
            help=help_text,
        )
    with col_btn:
        st.write("")  # align with text input
        st.write("")
        if st.button("Browse", key=f"browse_{state_key}", use_container_width=True):
            try:
                import tkinter  # noqa: F401
            except ImportError:
                st.warning(
                    "Folder picker needs tkinter (usually included with Python on Windows). "
                    "Paste the folder path manually instead."
                )
            else:
                selected = pick_folder(title=f"Select {label.lower()}")
                if selected:
                    st.session_state[pending_key] = selected
                    st.rerun()
    return (value or "").strip()


def preview_counts(source: Path) -> dict[str, int]:
    counts: dict[str, int] = {}
    for path in iter_source_files(source):
        category = categorize(path)
        counts[category] = counts.get(category, 0) + 1
    return counts


def render_counts(counts: dict[str, int]) -> None:
    if not counts:
        st.info("No files found in the source folder.")
        return
    cols = st.columns(3)
    for i, category in enumerate(CATEGORY_ORDER):
        if category in counts:
            cols[i % 3].metric(category, counts[category])
    st.caption(f"Total files: {sum(counts.values())}")


def validate_local_paths(source: Path, destination: Path) -> str | None:
    if not source.exists() or not source.is_dir():
        return f"Source folder not found: {source}"
    if source.resolve() == destination.resolve():
        return "Source and destination must be different folders."
    try:
        destination.resolve().relative_to(source.resolve())
        return "Destination cannot be inside the source folder."
    except ValueError:
        return None


def _sync_tree_checkbox_keys(tree: dict, selected: set[str]) -> None:
    """
    Align checkbox widget keys with selection set.

    Must only be called before those checkbox widgets are instantiated
    (e.g. in button handlers before st.rerun, or in on_change callbacks).
    """
    for path in all_paths(tree):
        st.session_state[f"tree_cb::{path}"] = path in selected


def _on_tree_checkbox_change(path: str) -> None:
    """Cascade parent/child selection safely via Streamlit on_change."""
    tree = st.session_state.get("iphone_tree")
    if not tree:
        return
    checked = bool(st.session_state.get(f"tree_cb::{path}", False))
    selected = set(st.session_state.get("tree_selected", []))
    updated = apply_check(selected, tree, path, checked)
    st.session_state["tree_selected"] = sorted(updated)
    # Safe here: on_change runs before widgets are instantiated
    _sync_tree_checkbox_keys(tree, updated)


def _render_tree_node(node: dict, depth: int = 0) -> None:
    path = node["path"]
    children = node.get("children") or []
    indent = "  " * depth  # em-space indent for tree levels
    child_count = len(children)
    suffix = f"  ({child_count} subfolders)" if child_count else ""
    icon = "📁" if depth == 0 or child_count else "📂"
    label = f"{indent}{icon} {node['name']}{suffix}"

    st.checkbox(
        label,
        key=f"tree_cb::{path}",
        on_change=_on_tree_checkbox_change,
        args=(path,),
    )

    for child in children:
        _render_tree_node(child, depth + 1)


def render_folder_tree_picker(
    tree: dict,
    device_name: str = "",
    *,
    allow_lazy_subfolders: bool = False,
) -> list[str]:
    """Checkbox tree with parent→child cascade selection."""
    st.markdown("#### iPhone folder tree")
    st.caption(
        "Check a parent folder to select its subfolders. "
        "Uncheck any folder you do not want to copy."
    )

    top_level = [c["path"] for c in (tree.get("children") or [])]
    remaining = [p for p in top_level if p not in st.session_state["completed_folders"]]
    st.caption(
        f"{count_folders(tree)} folders under **{tree.get('name', 'Internal Storage')}** · "
        f"{len(st.session_state['completed_folders'])} top-level completed · "
        f"{len(remaining)} top-level remaining"
    )

    batch_size = st.number_input(
        "Batch size (for Next batch)",
        min_value=1,
        max_value=max(1, len(top_level) or 1),
        value=int(st.session_state.get("batch_size", 5)),
        step=1,
        help="Select the next N remaining top-level folders (and their subfolders).",
    )
    st.session_state["batch_size"] = int(batch_size)

    b1, b2, b3, b4 = st.columns(4)
    if b1.button("Select all", use_container_width=True):
        selected = set(all_paths(tree))
        st.session_state["tree_selected"] = sorted(selected)
        _sync_tree_checkbox_keys(tree, selected)
        st.rerun()
    if b2.button("Next batch", use_container_width=True):
        batch = remaining[: int(batch_size)]
        selected = set()
        for path in batch:
            selected = apply_check(selected, tree, path, True)
        st.session_state["tree_selected"] = sorted(selected)
        _sync_tree_checkbox_keys(tree, selected)
        st.rerun()
    if b3.button("Clear selection", use_container_width=True):
        st.session_state["tree_selected"] = []
        _sync_tree_checkbox_keys(tree, set())
        st.rerun()
    if b4.button("Reset completed", use_container_width=True):
        st.session_state["completed_folders"] = []
        st.rerun()

    if allow_lazy_subfolders and device_name:
        st.caption("Lazy MTP subfolder loading is disabled in AFC mode.")

    # Prune invalid selections if tree reloaded (before checkbox widgets exist)
    valid = set(all_paths(tree))
    selected_now = {p for p in st.session_state.get("tree_selected", []) if p in valid}
    if sorted(selected_now) != list(st.session_state.get("tree_selected", [])):
        st.session_state["tree_selected"] = sorted(selected_now)
    # Keep widget keys in sync before instantiating checkboxes
    _sync_tree_checkbox_keys(tree, set(st.session_state.get("tree_selected", [])))

    with st.container(border=True):
        _render_tree_node(tree)

    selected_set = set(st.session_state.get("tree_selected", []))
    roots = selection_roots(selected_set)
    if roots:
        st.info(
            f"**{len(selected_set)}** folder(s) checked · "
            f"**{len(roots)}** copy root(s) will be used "
            f"(parent selection covers its subfolders)."
        )
    return roots


def render_copy_from_iphone() -> None:
    st.write(
        "Copy **real photo/video files** from iPhone to this laptop using Apple AFC "
        "(not Windows MTP — MTP often creates empty folders)."
    )

    if not is_windows():
        st.warning("This copy panel is intended for Windows with iTunes/Apple Devices installed.")

    st.info(
        "Windows Explorer MTP copy is unreliable (empty folders). "
        "This app now uses **Apple AFC** via pymobiledevice3 to pull actual files from DCIM."
    )

    top = st.columns([1, 1])
    with top[0]:
        refresh = st.button("Refresh devices", use_container_width=True)
    with top[1]:
        st.caption("Unlock iPhone, tap Trust, keep screen awake.")

    if refresh:
        st.session_state["iphone_tree"] = None
        st.session_state["tree_selected"] = []
        st.session_state["afc_serial"] = ""

    ok, detail = _pymobiledevice3_status()
    if not ok or not afc_available():
        st.error("Apple AFC support is not ready in the Python that is running this UI.")
        st.code(detail, language="text")
        st.write(f"This UI is running with: `{_python_exe()}`")
        st.write(f"Python version: `{sys.version.split()[0]}`")

        if sys.version_info >= (3, 13):
            st.warning(
                "Python 3.13 on Windows cannot install `lzfse` (needed by some "
                "iPhone tooling). Use a **Python 3.12** conda environment."
            )
            st.code(
                "\n".join(
                    [
                        "conda create -n iphone_copy python=3.12 -y",
                        "conda activate iphone_copy",
                        r"cd %USERPROFILE%\Documents\Automation_Projects\iphone_file_sorter",
                        "python -m pip install -r requirements.txt",
                        "python -m streamlit run app.py",
                    ]
                ),
                language="bash",
            )
        else:
            st.write("Install into **this exact Python** using the button below, or run:")
            st.code(
                f'"{_python_exe()}" -m pip install -U pymobiledevice3',
                language="bash",
            )
            if st.button("Install pymobiledevice3 into this Python", type="primary"):
                with st.spinner("Installing pymobiledevice3..."):
                    success, output = _install_pymobiledevice3()
                if success:
                    st.success(
                        "Install finished. Stop the app (Ctrl+C), start it again, then refresh."
                    )
                else:
                    st.error("Install failed. See output below.")
                st.code(output[-5000:] if output else "(no output)", language="text")

        st.caption(
            "Important: install and run Streamlit with the same conda environment. "
            "Python 3.12 is recommended on Windows."
        )
        return

    try:
        devices = list_afc_devices()
    except AfcError as exc:
        st.error(str(exc))
        return

    if not devices:
        st.warning(
            "No iPhone found over USB (AFC).\n\n"
            "- Unlock iPhone and tap **Trust**\n"
            "- Use a data cable\n"
            "- Keep **iTunes / Apple Devices** installed\n"
            "- Then click **Refresh devices**"
        )
        return

    labels = [f"{d['name']} ({d['serial']})" for d in devices]
    serials = [d["serial"] for d in devices]
    default_idx = 0
    if st.session_state.get("afc_serial") in serials:
        default_idx = serials.index(st.session_state["afc_serial"])

    choice = st.selectbox("iPhone (AFC)", options=labels, index=default_idx)
    serial = serials[labels.index(choice)]
    st.session_state["afc_serial"] = serial

    load = st.button("Load DCIM folder tree from iPhone", use_container_width=True)
    if load:
        try:
            with st.spinner("Reading DCIM folders via Apple AFC..."):
                tree = list_afc_folder_tree(serial, max_depth=2)
                st.session_state["iphone_tree"] = tree
                st.session_state["tree_selected"] = []
                _sync_tree_checkbox_keys(tree, set())
        except AfcError as exc:
            st.error(str(exc))
            st.session_state["iphone_tree"] = None

    tree = st.session_state.get("iphone_tree")
    if tree:
        st.success(
            f"Loaded **{tree.get('name', 'DCIM')}** "
            f"({count_folders(tree)} folders) via AFC."
        )
        selected_folders = render_folder_tree_picker(
            tree, allow_lazy_subfolders=False
        )
    else:
        selected_folders = []
        st.info("Click **Load DCIM folder tree from iPhone** to browse folders.")

    destination = path_with_browse(
        label="Paste to folder on laptop",
        state_key="copy_destination",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Copy",
        help_text="Files/folders from the iPhone will be copied here.",
    )

    also_sort = st.checkbox(
        "After copy, also sort into Images / Videos / Documents / Excel / PDF / Other",
        value=False,
    )
    sort_destination = ""
    if also_sort:
        sort_destination = path_with_browse(
            label="Sorted output folder",
            state_key="copy_sort_destination",
            placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
            help_text="Category folders will be created here after copying.",
        )

    st.caption(
        "AFC copies real files. Failed items are skipped and listed in an Excel log."
    )

    has_selection = bool(selected_folders)
    do_copy = st.button(
        "Copy selected folders to laptop",
        type="primary",
        use_container_width=True,
        disabled=not (has_selection and destination),
    )

    if not do_copy:
        return

    dest_path = Path(destination).expanduser()
    progress = st.progress(0, text="Starting AFC copy...")
    status = st.empty()
    st.info("Keep the iPhone unlocked. Files should appear in the destination as they download.")
    st.write(f"Destination: `{dest_path}`")

    def on_progress(name: str, index: int, total: int, stage: str) -> None:
        fraction = 0 if total == 0 else index / max(total, 1)
        progress.progress(
            min(max(fraction, 0.0), 1.0),
            text=f"{index}/{total}: {stage} · {name}",
        )
        status.write(f"**{stage.title()}:** `{name}`")

    try:
        with st.spinner("Copying from iPhone via Apple AFC (real files)..."):
            result = copy_afc_folders(
                serial,
                selected_folders,
                dest_path,
                on_progress=on_progress,
            )
    except AfcError as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Copy failed: {exc}")
        return

    completed = set(st.session_state.get("completed_folders", []))
    completed.update(selected_folders)
    st.session_state["completed_folders"] = sorted(completed)

    progress.progress(1.0, text="Batch finished")
    ok_n = len(result.succeeded)
    fail_n = len(result.failed)
    st.success(
        f"Batch complete: **{ok_n}** copied, **{fail_n}** skipped/failed."
    )
    st.write(f"Files saved under: `{dest_path.resolve()}`")

    if result.log_path and result.log_path.exists():
        st.write(f"Excel log: `{result.log_path.resolve()}`")
        try:
            data = result.log_path.read_bytes()
            st.download_button(
                "Download Excel log",
                data=data,
                file_name=result.log_path.name,
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        except OSError:
            pass

    if result.failed:
        st.subheader("Failed / skipped files")
        st.dataframe(result.failed, use_container_width=True)
    else:
        st.info("No failed files in this batch.")

    if also_sort:
        if not sort_destination.strip():
            st.error("Choose a sorted output folder.")
            return
        sorted_dest = Path(sort_destination).expanduser()
        err = validate_local_paths(dest_path, sorted_dest)
        if err:
            st.error(err)
            return
        with st.spinner("Sorting copied files..."):
            counts = sort_files(
                dest_path.resolve(),
                sorted_dest.resolve(),
                dry_run=False,
                move=False,
            )
        st.subheader("Sort summary")
        render_counts(counts)
        st.write(f"Sorted files are in: `{sorted_dest.resolve()}`")


def render_sort_local() -> None:
    st.write(
        "Sort files already on this laptop into "
        "**Images / Videos / Documents / Excel / PDF / Other**."
    )

    source_text = path_with_browse(
        label="Source folder (files to sort)",
        state_key="sort_source",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Copy",
        help_text="Folder on this PC that contains the files to organize.",
    )
    destination_text = path_with_browse(
        label="Destination folder (sorted output)",
        state_key="sort_destination",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
        help_text="Category folders will be created here.",
    )

    col_a, col_b = st.columns(2)
    with col_a:
        dry_run = st.checkbox("Dry run (preview only — no copy/move)", value=True)
    with col_b:
        move = st.checkbox("Move files instead of copy", value=False)

    if move and not dry_run:
        st.warning("Move mode will relocate files out of the source folder.")

    preview_btn, start_btn = st.columns(2)
    do_preview = preview_btn.button("Preview", use_container_width=True)
    do_start = start_btn.button(
        "Start sorting",
        type="primary",
        use_container_width=True,
    )

    source = Path(source_text).expanduser() if source_text else None
    destination = Path(destination_text).expanduser() if destination_text else None

    if do_preview:
        if source is None:
            st.error("Choose a source folder.")
            return
        if not source.exists() or not source.is_dir():
            st.error(f"Source folder not found: {source}")
            return
        with st.spinner("Scanning source folder..."):
            counts = preview_counts(source.resolve())
        st.subheader("Preview")
        render_counts(counts)

    if do_start:
        if source is None or destination is None:
            st.error("Choose both source and destination folders.")
            return

        error = validate_local_paths(source, destination)
        if error:
            st.error(error)
            return

        st.subheader("Progress")
        progress = st.progress(0, text="Starting...")
        log_box = st.empty()
        lines: list[str] = []

        def on_file(_src, _dest, category, index, total):
            fraction = 0 if total == 0 else index / total
            progress.progress(
                min(fraction, 1.0),
                text=f"{index} / {total} · {category}",
            )

        def log(message: str) -> None:
            lines.append(message)
            log_box.code("\n".join(lines[-40:]), language="text")

        try:
            with st.spinner("Sorting files..."):
                counts = sort_files(
                    source.resolve(),
                    destination.resolve(),
                    dry_run=dry_run,
                    move=move,
                    on_file=on_file,
                    log=log,
                )
        except (FileNotFoundError, OSError) as exc:
            st.error(str(exc))
            return

        progress.progress(1.0, text="Finished")
        st.subheader("Summary")
        if dry_run:
            st.info("Dry run complete — no files were changed.")
        else:
            st.success("Sorting complete.")
            st.write(f"Output folder: `{destination.resolve()}`")
        render_counts(counts)


def main() -> None:
    _init_state()

    st.title("iPhone File Copier & Sorter")
    st.write(
        "Two panels: **Copy** files from your iPhone to the laptop, "
        "then **Sort** them by file type."
    )

    with st.expander("Tips & supported file types", expanded=False):
        st.markdown(
            """
**Copy tips**
1. Unlock iPhone → tap **Trust this computer**
2. Turn off **Low Power Mode**, set **Auto-Lock** as long as possible
3. Keep the screen awake while copying
4. Copy in **batches** (use **Next batch**)
5. Failed files are **auto-skipped**; check the Excel log for details
6. If iCloud **Optimize iPhone Storage** is on, some photos may not be fully on-device

**Sort categories**
            """
        )
        for category, extensions in CATEGORY_EXTENSIONS.items():
            st.write(f"- **{category}**: {', '.join(sorted(extensions))}")

    copy_tab, sort_tab = st.tabs(
        [
            "1. Copy files",
            "2. Sort files",
        ]
    )

    with copy_tab:
        st.header("Copy files")
        st.caption("Copy from iPhone → choose destination folder on this laptop.")
        st.divider()
        render_copy_from_iphone()

    with sort_tab:
        st.header("Sort files")
        st.caption(
            "Organize files already on the laptop into Images, Videos, Documents, Excel, PDF, Other."
        )
        st.divider()
        render_sort_local()


if __name__ == "__main__":
    main()
