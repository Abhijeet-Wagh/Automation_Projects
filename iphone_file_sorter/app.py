#!/usr/bin/env python3
"""
iPhone File Copier & Sorter — Streamlit UI

Launch:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from folder_picker import pick_folder
from folder_tree import (
    all_paths,
    apply_check,
    count_folders,
    find_node,
    selection_roots,
)
from sort_files import CATEGORY_EXTENSIONS, categorize, iter_source_files, sort_files

try:
    from windows_iphone import (
        WindowsShellError,
        copy_folders_bulk,
        copy_folders_file_by_file,
        is_windows,
        list_iphone_child_folders,
        list_iphone_folder_tree,
        list_portable_apple_devices,
    )
except Exception:  # pragma: no cover - import safety on non-Windows
    WindowsShellError = RuntimeError  # type: ignore

    def is_windows() -> bool:
        return False

    def list_portable_apple_devices() -> list[str]:
        return []

    def list_iphone_folder_tree(_device_name: str, *, max_depth: int = 1):
        return {"name": "Internal Storage", "path": "", "children": []}

    def list_iphone_child_folders(_device_name: str, _parent_rel_path: str):
        return []

    def copy_folders_bulk(*_args, **_kwargs):
        raise RuntimeError("iPhone copy is only available on Windows.")

    def copy_folders_file_by_file(*_args, **_kwargs):
        raise RuntimeError("iPhone copy is only available on Windows.")


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


def render_folder_tree_picker(tree: dict, device_name: str) -> list[str]:
    """Checkbox tree with parent→child cascade selection."""
    st.markdown("#### iPhone folder tree")
    st.caption(
        "Top-level folders load first (fast). "
        "Check a parent to select its visible subfolders. "
        "Use **Load subfolders** only if you need deeper folders."
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

    # Lazy-load one level of subfolders for currently checked folders
    if st.button(
        "Load subfolders for checked folders",
        use_container_width=True,
        help=(
            "Optional. Scans only checked folders for subfolders. "
            "Can be slow on iPhone — prefer top-level selection when possible."
        ),
    ):
        selected_paths = [
            p for p in st.session_state.get("tree_selected", []) if p != ""
        ]
        if not selected_paths:
            st.warning("Check one or more folders first, then load their subfolders.")
        else:
            try:
                with st.spinner(
                    f"Loading subfolders for {len(selected_paths)} folder(s). "
                    "This may take a minute..."
                ):
                    for parent_path in selected_paths:
                        node = find_node(tree, parent_path)
                        if node is None:
                            continue
                        if node.get("children"):
                            continue  # already loaded
                        children = list_iphone_child_folders(device_name, parent_path)
                        node["children"] = children
                        # If parent is selected, auto-select new children
                        if parent_path in st.session_state.get("tree_selected", []):
                            selected = set(st.session_state["tree_selected"])
                            selected.update(c["path"] for c in children)
                            st.session_state["tree_selected"] = sorted(selected)
                    st.session_state["iphone_tree"] = tree
                    _sync_tree_checkbox_keys(
                        tree, set(st.session_state.get("tree_selected", []))
                    )
                st.success("Subfolder scan finished.")
                st.rerun()
            except WindowsShellError as exc:
                st.error(str(exc))

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
        "Copy files from a connected **iPhone** to a folder on this laptop. "
        "Select folders in the tree, choose where to paste, then start the copy."
    )

    if not is_windows():
        st.error("Direct iPhone copy is supported on **Windows** only.")
        return

    top = st.columns([1, 1])
    with top[0]:
        refresh = st.button("Refresh devices", use_container_width=True)
    with top[1]:
        st.caption("Unlock iPhone, tap Trust, keep screen awake.")

    try:
        devices = list_portable_apple_devices()
    except WindowsShellError as exc:
        st.error(str(exc))
        return

    if refresh:
        st.session_state["iphone_tree"] = None
        st.session_state["tree_selected"] = []

    if not devices:
        st.warning(
            "No iPhone/iPad found under **This PC**.\n\n"
            "- Unlock the iPhone and tap **Trust**\n"
            "- Use a data cable\n"
            "- Open File Explorer and confirm the phone appears\n"
            "- Then click **Refresh devices**"
        )
        return

    device = st.selectbox(
        "iPhone / iPad (copy from)",
        options=devices,
        index=devices.index(st.session_state["selected_device"])
        if st.session_state["selected_device"] in devices
        else 0,
    )
    st.session_state["selected_device"] = device

    load = st.button("Load folder tree from iPhone", use_container_width=True)
    if load:
        try:
            with st.spinner(
                "Reading top-level folders from iPhone (fast mode)..."
            ):
                # depth=1 only — deep scans hang on iPhone MTP photo folders
                tree = list_iphone_folder_tree(device, max_depth=1)
                st.session_state["iphone_tree"] = tree
                st.session_state["tree_selected"] = []
                _sync_tree_checkbox_keys(tree, set())
        except WindowsShellError as exc:
            st.error(str(exc))
            st.session_state["iphone_tree"] = None

    tree = st.session_state.get("iphone_tree")
    if tree:
        st.success(
            f"Loaded tree for **{tree.get('name', 'Internal Storage')}** "
            f"({count_folders(tree)} folders)."
        )
        selected_folders = render_folder_tree_picker(tree, device)
    else:
        selected_folders = []
        st.info("Click **Load folder tree from iPhone** to browse folders with checkboxes.")

    destination = path_with_browse(
        label="Paste to folder on laptop",
        state_key="copy_destination",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Copy",
        help_text="Files/folders from the iPhone will be copied here.",
    )

    copy_mode = st.radio(
        "Copy mode",
        options=[
            "Fast: whole folders (Windows copy dialog) — recommended",
            "Detailed: file-by-file (slower, auto-skip + Excel log)",
        ],
        index=0,
        help=(
            "Fast mode copies each selected folder like File Explorer and shows the "
            "Windows progress window. File-by-file mode is much slower on iPhone."
        ),
    )
    use_file_by_file = copy_mode.startswith("Detailed")

    file_timeout = 45
    if use_file_by_file:
        file_timeout = st.slider(
            "Per-file wait timeout (seconds)",
            min_value=15,
            max_value=180,
            value=30,
            help="If a file doesn’t finish in time, it is skipped and logged.",
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

    if use_file_by_file:
        st.caption(
            "File-by-file can look stuck on file 1/N for up to the timeout. "
            "Keep the iPhone unlocked. Prefer Fast mode for large folders."
        )
    else:
        st.caption(
            "A Windows copy window should appear. If a file errors, click Skip and "
            "let the rest continue. An Excel log of arrived files is saved at the end."
        )

    # selected_folders can be [""] for full root — treat as valid selection
    has_selection = bool(selected_folders) or selected_folders == [""]
    # selection_roots returns [""] when root checked — bool([""]) is True. Good.
    # when nothing selected, roots is []. Good.

    do_copy = st.button(
        "Copy selected folders to laptop",
        type="primary",
        use_container_width=True,
        disabled=not (has_selection and destination),
    )

    if not do_copy:
        return

    dest_path = Path(destination).expanduser()
    progress = st.progress(0, text="Starting copy...")
    status = st.empty()
    tip = st.empty()
    tip.info(
        "Keep the iPhone unlocked. "
        + (
            "Look behind the browser / on the taskbar for a Windows copy window. "
            "Files should appear under your destination folder as it runs."
            if not use_file_by_file
            else f"Waiting up to {file_timeout}s per file, then auto-skip."
        )
    )
    st.write(f"Destination: `{dest_path}`")

    def on_progress(name: str, index: int, total: int, stage: str) -> None:
        fraction = 0 if total == 0 else index / max(total, 1)
        # For waiting stage, show partial progress within the current item
        progress.progress(
            min(max(fraction, 0.0), 1.0),
            text=f"{index}/{total}: {stage} · {name}",
        )
        status.write(f"**{stage.title()}:** `{name}`")

    try:
        if use_file_by_file:
            with st.spinner("Copying file-by-file from iPhone..."):
                result = copy_folders_file_by_file(
                    device,
                    selected_folders,
                    dest_path,
                    file_timeout_sec=float(file_timeout),
                    on_progress=on_progress,
                )
        else:
            with st.spinner(
                "Copying whole folders via Windows dialog (this can take a while)..."
            ):
                result = copy_folders_bulk(
                    device,
                    selected_folders,
                    dest_path,
                    on_progress=on_progress,
                )
    except WindowsShellError as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Copy failed: {exc}")
        return

    # Mark top-level folders from this selection as completed
    completed = set(st.session_state.get("completed_folders", []))
    for path in selected_folders:
        top = path.split("/", 1)[0] if path else ""
        if top:
            completed.add(top)
        else:
            # full root copy — mark all top-level children completed
            for child in tree.get("children") or []:
                completed.add(child["path"])
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
