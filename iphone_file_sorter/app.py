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
                    st.session_state[state_key] = selected
                    st.rerun()
    return value.strip() if value else ""


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


def _render_tree_node(node: dict, tree_root: dict, depth: int = 0) -> None:
    path = node["path"]
    children = node.get("children") or []
    selected = set(st.session_state.get("tree_selected", []))
    checked = path in selected
    indent = "  " * depth  # em-space indent for tree levels
    child_count = len(children)
    suffix = f"  ({child_count} subfolders)" if child_count else ""
    icon = "📁" if depth == 0 or child_count else "📂"
    label = f"{indent}{icon} {node['name']}{suffix}"

    key = f"tree_cb::{path}"
    # Initialize only if missing — never overwrite after widget exists
    if key not in st.session_state:
        st.session_state[key] = checked

    st.checkbox(
        label,
        key=key,
        on_change=_on_tree_checkbox_change,
        args=(path,),
    )

    for child in children:
        _render_tree_node(child, tree_root, depth + 1)


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
        _render_tree_node(tree, tree)

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
    st.subheader("Copy from iPhone → laptop")
    st.write(
        "Select folders from the iPhone tree (parent check selects subfolders), "
        "copy file-by-file with auto-skip, and get an Excel failure log."
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

    file_timeout = st.slider(
        "Per-file wait timeout (seconds)",
        min_value=15,
        max_value=300,
        value=90,
        help="If a file doesn’t finish copying in time, it is logged as failed and skipped.",
    )

    also_sort = st.checkbox(
        "After copy, also sort into Images / Videos / Documents / Excel / PDF / Other",
        value=False,
    )
    sort_destination = ""
    if also_sort:
        sort_destination = path_with_browse(
            label="Sorted output folder",
            state_key="sort_destination",
            placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
            help_text="Category folders will be created here after copying.",
        )

    st.caption(
        "Failed files are skipped automatically. An Excel log is saved in the destination "
        "folder with source path, destination path, and error details."
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

    def on_progress(name: str, index: int, total: int, stage: str) -> None:
        fraction = 0 if total == 0 else index / total
        progress.progress(
            min(max(fraction, 0.0), 1.0),
            text=f"{index}/{total}: {stage} · {name}",
        )
        status.write(f"**{stage.title()}:** `{name}`")

    try:
        with st.spinner("Copying files from iPhone (skipping failures)..."):
            result = copy_folders_file_by_file(
                device,
                selected_folders,
                dest_path,
                file_timeout_sec=float(file_timeout),
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
    st.subheader("Sort files already on this laptop")
    st.write(
        "Use this after files are already copied locally. "
        "Browse the source folder and a destination for sorted categories."
    )

    source_text = path_with_browse(
        label="Source folder (copy from)",
        state_key="sort_source",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Copy",
        help_text="Folder that already contains the iPhone files on this PC.",
    )
    destination_text = path_with_browse(
        label="Destination folder (paste / sort into)",
        state_key="sort_destination",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
        help_text="Sorted category folders will be created here.",
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
        "Copy folders from your connected iPhone to this laptop, "
        "then optionally sort them by file type."
    )

    with st.expander("Tips for reliable iPhone copy", expanded=False):
        st.markdown(
            """
1. Unlock iPhone → tap **Trust this computer**
2. Turn off **Low Power Mode**, set **Auto-Lock** as long as possible
3. Keep the screen awake while copying
4. Copy in **batches** (use **Next batch**)
5. Failed files are **auto-skipped**; check the Excel log for details
6. If iCloud **Optimize iPhone Storage** is on, some photos may not be fully on-device
            """
        )
        st.markdown("**Sort categories**")
        for category, extensions in CATEGORY_EXTENSIONS.items():
            st.write(f"- **{category}**: {', '.join(sorted(extensions))}")

    mode = st.radio(
        "What do you want to do?",
        options=[
            "Copy from iPhone to laptop",
            "Sort files already on laptop",
        ],
        horizontal=False,
    )

    st.divider()
    if mode == "Copy from iPhone to laptop":
        render_copy_from_iphone()
    else:
        render_sort_local()


if __name__ == "__main__":
    main()
