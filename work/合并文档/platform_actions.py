"""Small operating-system adapters for the shared V5 interface."""

from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from pathlib import Path


def set_window_icon(root: tk.Misc, assets: Path) -> None:
    if sys.platform == "darwin":
        icon = tk.PhotoImage(master=root, file=str(assets / "a-logo.png"))
        try:
            root.iconphoto(True, icon)
        except tk.TclError:
            pass  # The .app bundle icon still supplies the macOS Dock icon.
        root._merge_app_icon = icon  # Keep the Tcl image alive for the window.
    else:
        root.iconbitmap(str(assets / "app_icon.ico"))


def reveal_file(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)])
    else:
        subprocess.Popen(["explorer.exe", "/select,", str(path)])
