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

from content_categories import (
    CATEGORIES,
    categories_by_group,
    default_selected_ids,
)
from folder_picker import pick_folder
from folder_tree import all_paths, apply_check, count_folders, selection_roots
from sort_files import CATEGORY_EXTENSIONS, categorize, iter_source_files, sort_files

_AFC_IMPORT_ERROR = ""
try:
    from afc_iphone import (
        AfcError,
        afc_available,
        copy_afc_folders,
        copy_content_categories,
        list_afc_devices,
        list_afc_folder_tree,
        summarize_copy_rows,
    )
except Exception as exc:  # pragma: no cover
    _AFC_IMPORT_ERROR = f"{type(exc).__name__}: {exc}"
    AfcError = RuntimeError  # type: ignore

    def afc_available() -> bool:
        return False

    def list_afc_devices():
        return []

    def list_afc_folder_tree(_serial: str, *, max_depth: int = 2):
        return {"name": "iPhone", "path": "", "children": []}

    def copy_afc_folders(*_args, **_kwargs):
        raise RuntimeError("AFC copy unavailable.")

    def copy_content_categories(*_args, **_kwargs):
        raise RuntimeError("AFC copy unavailable.")

    def summarize_copy_rows(succeeded, failed):
        return {
            "copied": len(succeeded or []),
            "already_on_disk": 0,
            "system_skipped": 0,
            "timeout": 0,
            "other_failed": len(failed or []),
        }


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
        "tree_expanded": [],
        "tree_expand_initialized": False,
        "completed_folders": [],
        "batch_size": 5,
        "content_selected": default_selected_ids(),
        "backup_password": "",
        "copy_mode": "categories",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    for cat in CATEGORIES:
        ck = f"cat::{cat.id}"
        if ck not in st.session_state:
            st.session_state[ck] = bool(cat.default_on)


def _inject_tree_scroll_styles() -> None:
    """Force a visible vertical scrollbar on the folder-tree box.

    Streamlit uses overflow:auto, and Windows often hides overlay scrollbars
    until you hover/scroll. This CSS keeps a clear scrollbar track + thumb.
    """
    st.markdown(
        """
<style>
/* Target Streamlit keyed container (key="iphone_folder_tree") */
div.st-key-iphone_folder_tree,
div[class*="st-key-iphone_folder_tree"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(.iphone-tree-scroll-marker) {
  height: 420px !important;
  max-height: 420px !important;
  overflow-y: scroll !important;
  overflow-x: hidden !important;
  scrollbar-gutter: stable;
  scrollbar-width: auto; /* Firefox */
  scrollbar-color: #4b5563 #d1d5db;
}

/* Chromium / Edge — always show a thick, obvious scrollbar */
div.st-key-iphone_folder_tree::-webkit-scrollbar,
div[class*="st-key-iphone_folder_tree"]::-webkit-scrollbar,
div[data-testid="stVerticalBlockBorderWrapper"]:has(.iphone-tree-scroll-marker)::-webkit-scrollbar {
  width: 14px;
  -webkit-appearance: none;
}
div.st-key-iphone_folder_tree::-webkit-scrollbar-track,
div[class*="st-key-iphone_folder_tree"]::-webkit-scrollbar-track,
div[data-testid="stVerticalBlockBorderWrapper"]:has(.iphone-tree-scroll-marker)::-webkit-scrollbar-track {
  background: #d1d5db;
}
div.st-key-iphone_folder_tree::-webkit-scrollbar-thumb,
div[class*="st-key-iphone_folder_tree"]::-webkit-scrollbar-thumb,
div[data-testid="stVerticalBlockBorderWrapper"]:has(.iphone-tree-scroll-marker)::-webkit-scrollbar-thumb {
  background: #4b5563;
  border-radius: 7px;
  border: 2px solid #d1d5db;
  min-height: 40px;
}
</style>
        """,
        unsafe_allow_html=True,
    )


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


def _expand_key(path: str) -> str:
    return f"tree_exp::{path if path else '__root__'}"


def _is_expanded(path: str) -> bool:
    return path in set(st.session_state.get("tree_expanded", []))


def _toggle_expanded(path: str) -> None:
    expanded = set(st.session_state.get("tree_expanded", []))
    if path in expanded:
        expanded.discard(path)
    else:
        expanded.add(path)
    st.session_state["tree_expanded"] = sorted(expanded)


def _set_expanded_paths(paths: list[str]) -> None:
    st.session_state["tree_expanded"] = sorted(set(paths))


def _render_tree_node(node: dict, depth: int = 0) -> None:
    """Render one tree row with expand/collapse arrow + checkbox."""
    path = node["path"]
    children = node.get("children") or []
    has_children = bool(children)
    is_open = _is_expanded(path)

    child_count = len(children)
    suffix = f"  ({child_count} subfolders)" if child_count else ""
    icon = "📁" if has_children or depth == 0 else "📂"
    label = f"{icon} {node['name']}{suffix}"

    # Indent | arrow | checkbox
    indent_weight = max(0.2, depth * 0.35)
    arrow_weight = 0.45
    check_weight = max(3.5, 6.0 - indent_weight)
    c_indent, c_arrow, c_check = st.columns([indent_weight, arrow_weight, check_weight])

    with c_indent:
        st.write("")

    with c_arrow:
        if has_children:
            arrow = "▼" if is_open else "▶"
            if st.button(
                arrow,
                key=_expand_key(path),
                help="Expand / collapse folder",
                use_container_width=True,
            ):
                _toggle_expanded(path)
                st.rerun()
        else:
            # Leaf marker (no expand arrow)
            st.markdown(
                "<div style='text-align:center; opacity:0.45;'>•</div>",
                unsafe_allow_html=True,
            )

    with c_check:
        st.checkbox(
            label,
            key=f"tree_cb::{path}",
            on_change=_on_tree_checkbox_change,
            args=(path,),
        )

    if has_children and is_open:
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
        "Use ▶ / ▼ to expand or collapse folders. "
        "Check a parent folder to select its subfolders."
    )

    top_level = [c["path"] for c in (tree.get("children") or [])]
    remaining = [p for p in top_level if p not in st.session_state["completed_folders"]]
    st.caption(
        f"{count_folders(tree)} folders under **{tree.get('name', 'Internal Storage')}** · "
        f"{len(st.session_state['completed_folders'])} top-level completed · "
        f"{len(remaining)} top-level remaining"
    )

    batch_max = max(1, len(top_level) or 1)
    batch_default = min(int(st.session_state.get("batch_size", 5)), batch_max)
    batch_default = max(1, batch_default)
    # Keep session value in range before the widget is created (Streamlit
    # raises if value > max_value, e.g. old default 5 with only 2 top folders).
    st.session_state["batch_size"] = batch_default
    batch_size = st.number_input(
        "Batch size (for Next batch)",
        min_value=1,
        max_value=batch_max,
        value=batch_default,
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

    e1, e2 = st.columns(2)
    if e1.button("Expand all", use_container_width=True):
        _set_expanded_paths(all_paths(tree))
        st.rerun()
    if e2.button("Collapse all", use_container_width=True):
        # Keep only the root visible/collapsed-closed
        _set_expanded_paths([])
        st.rerun()

    if allow_lazy_subfolders and device_name:
        st.caption("Lazy MTP subfolder loading is disabled in AFC mode.")

    # Prune invalid selections / expanded paths if tree reloaded
    valid = set(all_paths(tree))
    selected_now = {p for p in st.session_state.get("tree_selected", []) if p in valid}
    if sorted(selected_now) != list(st.session_state.get("tree_selected", [])):
        st.session_state["tree_selected"] = sorted(selected_now)

    expanded_now = [p for p in st.session_state.get("tree_expanded", []) if p in valid]
    # Expand root only once when a tree is first loaded — never force-reopen
    # after the user collapses DCIM / uses Collapse all.
    if not st.session_state.get("tree_expand_initialized", False):
        if tree.get("path") in valid:
            expanded_now = [tree["path"]]
        st.session_state["tree_expand_initialized"] = True
    st.session_state["tree_expanded"] = expanded_now

    # Keep widget keys in sync before instantiating checkboxes
    _sync_tree_checkbox_keys(tree, set(st.session_state.get("tree_selected", [])))

    st.caption(
        "Folder list is inside the box below — use the **vertical scrollbar on the right** "
        "when the tree is taller than the box."
    )
    _inject_tree_scroll_styles()
    # Fixed-height bordered box + CSS force a visible vertical scrollbar
    try:
        tree_box = st.container(
            border=True,
            height=420,
            key="iphone_folder_tree",
            autoscroll=False,
        )
    except TypeError:
        # Older Streamlit: height/key/autoscroll may be missing
        try:
            tree_box = st.container(border=True, height=420)
        except TypeError:
            tree_box = st.container()

    with tree_box:
        st.markdown(
            '<div class="iphone-tree-scroll-marker" style="display:none"></div>',
            unsafe_allow_html=True,
        )
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
        "Copy the content you care about from iPhone → laptop: "
        "camera media, WhatsApp media, documents, recordings, and optionally "
        "SMS/Contacts/Notes/chat databases via selective backup."
    )

    if not is_windows():
        st.warning("This copy panel is intended for Windows with iTunes/Apple Devices installed.")

    st.info(
        "Recommended mode uses **content categories** with file-type filters "
        "(images, videos, audio, PDFs/docs). PhotoData system junk is excluded. "
        "Chat **text**, Contacts, SMS, and Notes use a selective backup (database files)."
    )

    top = st.columns([1, 1])
    with top[0]:
        refresh = st.button("Refresh devices", use_container_width=True)
    with top[1]:
        st.caption("Unlock iPhone, tap Trust, keep screen awake.")

    if refresh:
        st.session_state["iphone_tree"] = None
        st.session_state["tree_selected"] = []
        st.session_state["tree_expanded"] = []
        st.session_state["tree_expand_initialized"] = False
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

    st.markdown("#### What do you want to copy?")
    st.caption(
        "Choose content types below. The app **filters to useful files** "
        "(images, videos, audio, PDFs/docs) and skips PhotoData caches / system junk. "
        "SMS, Contacts, Notes, and WhatsApp **chat text** need a selective iPhone backup."
    )

    mode = st.radio(
        "Copy mode",
        options=["categories", "advanced_tree"],
        format_func=lambda m: (
            "Recommended: content categories (filtered)"
            if m == "categories"
            else "Advanced: raw folder tree"
        ),
        horizontal=True,
        key="copy_mode",
    )

    selected_folders: list[str] = []
    selected_categories: list[str] = []
    backup_password = ""

    if mode == "categories":
        groups = categories_by_group()
        group_titles = {
            "media": "Photos, videos, audio, documents, WhatsApp media",
            "messages": "Chats / SMS (selective backup — databases, not pretty HTML)",
            "people": "Contacts, call history, Notes (selective backup)",
        }
        chosen: list[str] = []
        for group_key, cats in groups.items():
            if not cats:
                continue
            st.markdown(f"**{group_titles.get(group_key, group_key)}**")
            for cat in cats:
                checked = st.checkbox(
                    cat.label,
                    key=f"cat::{cat.id}",
                    help=cat.description,
                )
                if checked:
                    chosen.append(cat.id)
        st.session_state["content_selected"] = chosen
        selected_categories = chosen

        needs_backup = any(
            c.kind == "backup"
            for c in CATEGORIES
            if c.id in selected_categories
        )
        if needs_backup:
            st.warning(
                "**Backup can look stuck:** Apple’s backup protocol is slow. The bar may sit "
                "at ~40% for a long time while files still appear on disk. "
                "Empty `00`–`ff` folders are normal. "
                "Watch `iPhone_Backup_Selected\\_backup_heartbeat.txt` — if that file’s "
                "timestamp and `files_on_disk` keep updating, it is still working. "
                "Only stop (Ctrl+C) if the heartbeat file is unchanged for **15+ minutes** "
                "and Explorer shows no new/changed files."
            )
            backup_password = st.text_input(
                "iPhone backup password (only if encrypted backups are enabled)",
                type="password",
                key="backup_password",
                help=(
                    "Required when the iPhone encrypts local backups. "
                    "Set/check this in Finder or iTunes backup settings."
                ),
            )
            st.caption(
                "Backup items are saved under `iPhone_Backup_Selected` in your destination. "
                "They are database files — use a viewer/exporter later for readable chats."
            )

        c1, c2 = st.columns(2)
        if c1.button("Select recommended", use_container_width=True):
            recommended = set(default_selected_ids())
            for cat in CATEGORIES:
                st.session_state[f"cat::{cat.id}"] = cat.id in recommended
            st.session_state["content_selected"] = default_selected_ids()
            st.rerun()
        if c2.button("Clear categories", use_container_width=True):
            for cat in CATEGORIES:
                st.session_state[f"cat::{cat.id}"] = False
            st.session_state["content_selected"] = []
            st.rerun()

        if selected_categories:
            st.info(
                f"**{len(selected_categories)}** content type(s) selected. "
                "Junk/system files are excluded automatically while copying."
            )
    else:
        load = st.button("Load folder tree from iPhone", use_container_width=True)
        if load:
            try:
                with st.spinner("Reading Media + Apps folders via Apple AFC..."):
                    tree = list_afc_folder_tree(serial, max_depth=3)
                    st.session_state["iphone_tree"] = tree
                    st.session_state["tree_selected"] = []
                    expanded = [tree.get("path", "")]
                    for child in tree.get("children") or []:
                        if child.get("path") == "media":
                            expanded.append(child["path"])
                    st.session_state["tree_expanded"] = expanded
                    st.session_state["tree_expand_initialized"] = True
                    _sync_tree_checkbox_keys(tree, set())
            except AfcError as exc:
                st.error(str(exc))
                st.session_state["iphone_tree"] = None
                st.session_state["tree_expand_initialized"] = False

        tree = st.session_state.get("iphone_tree")
        if tree:
            st.warning(
                "Advanced mode copies whole folders. Prefer **Media → DCIM** only; "
                "avoid PhotoData."
            )
            selected_folders = render_folder_tree_picker(
                tree, allow_lazy_subfolders=False
            )
        else:
            st.info("Click **Load folder tree from iPhone** for advanced browsing.")

    destination = path_with_browse(
        label="Paste to folder on laptop",
        state_key="copy_destination",
        placeholder=r"C:\iPhone_Copy",
        help_text="Prefer a local folder (not OneDrive).",
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
            placeholder=r"C:\iPhone_Sorted",
            help_text="Category folders will be created here after copying.",
        )

    st.caption(
        "Copy runs file-by-file with timeouts. Re-run to resume "
        "(already-copied files are skipped). Excel log lists real failures."
    )

    has_selection = bool(selected_categories) if mode == "categories" else bool(selected_folders)
    do_copy = st.button(
        "Copy selected content to laptop",
        type="primary",
        use_container_width=True,
        disabled=not (has_selection and destination),
    )

    if not do_copy:
        return

    dest_path = Path(destination).expanduser()
    if "onedrive" in str(dest_path).lower():
        st.warning(
            "Destination is under **OneDrive**. Cloud sync can make copies look frozen "
            "or very slow. Prefer a local folder such as "
            r"`C:\Users\abhij\Documents\iPhone_Copy` (outside OneDrive) or `C:\iPhone_Copy`."
        )
    progress = st.progress(0, text="Starting copy...")
    status = st.empty()
    st.info(
        "Keep the iPhone unlocked. Filtered copy starts immediately — "
        "watch this status line and the destination folder."
    )
    st.write(f"Destination: `{dest_path}`")

    def on_progress(name: str, index: int, total: int, stage: str) -> None:
        fraction = 0 if total == 0 else index / max(total, 1)
        progress.progress(
            min(max(fraction, 0.0), 0.99),
            text=f"{stage} · {index}/{total} · {name}",
        )
        status.write(f"**{stage.title()}** ({index}/{total}): `{name}`")

    try:
        with st.spinner(
            "Copying from iPhone… if this is a backup, progress may pause for a long time "
            "while files still write to disk. Check _backup_heartbeat.txt in the destination."
        ):
            if mode == "categories":
                result = copy_content_categories(
                    serial,
                    selected_categories,
                    dest_path,
                    backup_password=backup_password or "",
                    on_progress=on_progress,
                )
            else:
                result = copy_afc_folders(
                    serial,
                    selected_folders,
                    dest_path,
                    on_progress=on_progress,
                )
    except AfcError as exc:
        detail = str(exc).strip() or repr(exc)
        st.error(detail)
        st.caption(
            "If WhatsApp media fails, iOS may block app USB access — "
            "use WhatsApp Export, or enable the WhatsApp chat database backup option. "
            "Avoid OneDrive destinations."
        )
        return
    except Exception as exc:  # noqa: BLE001
        detail = str(exc).strip() or repr(exc)
        st.error(f"Copy failed: {detail}")
        return

    if mode == "advanced_tree":
        completed = set(st.session_state.get("completed_folders", []))
        completed.update(selected_folders)
        st.session_state["completed_folders"] = sorted(completed)

    progress.progress(1.0, text="Batch finished")
    fail_n = len(result.failed)
    breakdown = summarize_copy_rows(result.succeeded, result.failed)
    filtered_n = int(getattr(result, "filtered_out", 0) or 0)
    st.success(
        f"Batch complete: **{breakdown['copied']}** copied, "
        f"**{breakdown['already_on_disk']}** already on disk, "
        f"**{fail_n}** failed/logged skips, "
        f"**{filtered_n}** junk/non-matching files ignored."
    )
    st.write(f"Files saved under: `{dest_path.resolve()}`")

    if filtered_n or breakdown["system_skipped"] or fail_n:
        st.info(
            "Ignored/filtered files are mostly system junk or wrong type "
            "(not images/videos/docs/audio). "
            f"Logged failures: **{breakdown['timeout']}** timeouts, "
            f"**{breakdown['other_failed']}** other errors."
        )

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
        st.caption(
            "Open the Excel log and filter the Error column — "
            "`Skipped system cache/DB file` means intentional (not a photo)."
        )
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
