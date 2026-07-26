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
from sort_files import CATEGORY_EXTENSIONS, categorize, iter_source_files, sort_files

try:
    from windows_iphone import (
        WindowsShellError,
        copy_named_items_to_folder,
        is_windows,
        list_iphone_media_folders,
        list_portable_apple_devices,
    )
except Exception:  # pragma: no cover - import safety on non-Windows
    WindowsShellError = RuntimeError  # type: ignore

    def is_windows() -> bool:
        return False

    def list_portable_apple_devices() -> list[str]:
        return []

    def list_iphone_media_folders(_device_name: str) -> list[str]:
        return []

    def copy_named_items_to_folder(*_args, **_kwargs):
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
        "iphone_folders": [],
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


def render_copy_from_iphone() -> None:
    st.subheader("Copy from iPhone → laptop")
    st.write(
        "Select the connected iPhone, choose folders to copy, and pick where to paste "
        "them on this laptop."
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
        st.session_state["iphone_folders"] = []

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

    load = st.button("Load folders from iPhone", use_container_width=True)
    if load:
        try:
            with st.spinner("Reading folders from iPhone..."):
                st.session_state["iphone_folders"] = list_iphone_media_folders(device)
        except WindowsShellError as exc:
            st.error(str(exc))
            st.session_state["iphone_folders"] = []

    folders = st.session_state.get("iphone_folders") or []
    if folders:
        st.success(f"Found {len(folders)} folders on the iPhone.")
        selected_folders = st.multiselect(
            "Folders to copy",
            options=folders,
            default=folders,
            help="Tip: if Explorer copy errors occur, select fewer folders at a time.",
        )
    else:
        selected_folders = []
        st.info("Click **Load folders from iPhone** to list Internal Storage folders.")

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
            state_key="sort_destination",
            placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
            help_text="Category folders will be created here after copying.",
        )

    st.caption(
        "Windows may show its normal copy dialog. If you see "
        "“The requested value cannot be determined”, skip the bad file and continue, "
        "or copy fewer folders. Keep the iPhone unlocked."
    )

    do_copy = st.button(
        "Copy to laptop",
        type="primary",
        use_container_width=True,
        disabled=not (selected_folders and destination),
    )

    if not do_copy:
        return

    dest_path = Path(destination).expanduser()
    progress = st.progress(0, text="Starting copy...")
    status = st.empty()

    def on_progress(name: str, index: int, total: int, stage: str) -> None:
        fraction = 0 if total == 0 else (index - (0 if stage == "done" else 1)) / total
        if stage == "done":
            fraction = index / total
        progress.progress(
            min(max(fraction, 0.0), 1.0),
            text=f"{index}/{total}: {name} ({stage})",
        )
        status.write(f"{stage.title()}: `{name}`")

    try:
        with st.spinner("Copying from iPhone via Windows Shell..."):
            copied = copy_named_items_to_folder(
                device,
                selected_folders,
                dest_path,
                on_progress=on_progress,
                silent=False,
            )
    except WindowsShellError as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"Copy failed: {exc}")
        return

    progress.progress(1.0, text="Copy finished")
    st.success(f"Copied {len(copied)} item(s) to `{dest_path.resolve()}`")

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
4. Prefer copying **a few folders at a time**
5. If a file errors with *“The requested value cannot be determined”*, skip it and continue
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
