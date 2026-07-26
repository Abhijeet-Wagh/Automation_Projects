#!/usr/bin/env python3
"""
iPhone File Sorter — Streamlit UI

Launch:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from sort_files import CATEGORY_EXTENSIONS, categorize, iter_source_files, sort_files


st.set_page_config(
    page_title="iPhone File Sorter",
    page_icon="📁",
    layout="centered",
)

CATEGORY_ORDER = ("Images", "Videos", "Documents", "Excel", "PDF", "Other")


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


def validate_paths(source: Path, destination: Path) -> str | None:
    if not source.exists() or not source.is_dir():
        return f"Source folder not found: {source}"
    if source.resolve() == destination.resolve():
        return "Source and destination must be different folders."
    try:
        destination.resolve().relative_to(source.resolve())
        return "Destination cannot be inside the source folder."
    except ValueError:
        return None


def main() -> None:
    st.title("iPhone File Sorter")
    st.write(
        "Sort files copied from your iPhone **Internal Storage** into "
        "Images, Videos, Documents, Excel, PDF, and Other."
    )

    with st.expander("How to use", expanded=False):
        st.markdown(
            """
1. In File Explorer, copy folders from  
   `This PC → Apple iPhone → Internal Storage` to a local folder.
2. Paste that local path as **Source folder** below.
3. Choose a **Destination folder** for sorted output.
4. Click **Preview**, then **Start sorting**.
5. Keep **Dry run** on first to preview safely.
            """
        )
        st.markdown("**Supported types**")
        for category, extensions in CATEGORY_EXTENSIONS.items():
            st.write(f"- **{category}**: {', '.join(sorted(extensions))}")

    st.subheader("Folders")
    source_text = st.text_input(
        "Source folder",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Copy",
        help="Folder that contains the copied iPhone folders (202403_b, etc.)",
    )
    destination_text = st.text_input(
        "Destination folder",
        placeholder=r"C:\Users\YourName\Documents\iPhone_Sorted",
        help="Sorted category folders will be created here.",
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

    source = Path(source_text.strip()).expanduser() if source_text.strip() else None
    destination = (
        Path(destination_text.strip()).expanduser() if destination_text.strip() else None
    )

    if do_preview:
        if source is None:
            st.error("Enter a source folder path.")
            st.stop()
        if not source.exists() or not source.is_dir():
            st.error(f"Source folder not found: {source}")
            st.stop()
        with st.spinner("Scanning source folder..."):
            counts = preview_counts(source.resolve())
        st.subheader("Preview")
        render_counts(counts)

    if do_start:
        if source is None:
            st.error("Enter a source folder path.")
            st.stop()
        if destination is None:
            st.error("Enter a destination folder path.")
            st.stop()

        error = validate_paths(source, destination)
        if error:
            st.error(error)
            st.stop()

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
            # Keep the on-screen log readable
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
            st.stop()

        progress.progress(1.0, text="Finished")
        st.subheader("Summary")
        if dry_run:
            st.info("Dry run complete — no files were changed.")
        else:
            st.success("Sorting complete.")
            st.write(f"Output folder: `{destination.resolve()}`")
        render_counts(counts)


if __name__ == "__main__":
    main()
