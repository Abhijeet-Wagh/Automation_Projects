"""Local folder picker using the system dialog (tkinter)."""

from __future__ import annotations


def pick_folder(title: str = "Select folder") -> str | None:
    """
    Open a native folder-selection dialog.

    Returns the selected path, or None if cancelled / unavailable.
    On Windows this can usually browse This PC; MTP/iPhone folders may or
    may not return a normal filesystem path depending on the OS.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    root.update()

    selected = filedialog.askdirectory(title=title, mustexist=True)
    root.destroy()
    return selected or None
