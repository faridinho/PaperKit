from __future__ import annotations

import os
import sys
import webbrowser
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox




# Optional drag-and-drop support.
# Install with: pip install tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES
    DND_AVAILABLE = True
except Exception:
    DND_FILES = None
    DND_AVAILABLE = False


def _first_drop_path(widget, event) -> str:
    """Return the first path from a tkinterdnd2 drop event."""
    try:
        paths = widget.tk.splitlist(event.data)
    except Exception:
        paths = [event.data]
    if not paths:
        return ""
    return str(paths[0]).strip().strip("{}")


def _register_drop_target(widget, handler) -> bool:
    """Register a drop target on a CTk widget and its inner Tk widget if present."""
    if not DND_AVAILABLE:
        return False

    targets = [widget]
    for attr in ("_entry", "_textbox", "_canvas"):
        inner = getattr(widget, attr, None)
        if inner is not None:
            targets.append(inner)

    ok = False
    for target in targets:
        try:
            target.drop_target_register(DND_FILES)
            target.dnd_bind("<<Drop>>", handler)
            ok = True
        except Exception:
            pass
    return ok


def enable_folder_drop(widget, variable: ctk.StringVar, on_change=None) -> bool:
    """Enable dropping a folder onto an entry. If a file is dropped, use its parent folder."""
    def handle(event):
        raw = _first_drop_path(widget, event)
        if not raw:
            return
        path = Path(raw)
        if path.is_file():
            path = path.parent
        variable.set(str(path))
        if on_change:
            on_change()

    return _register_drop_target(widget, handle)


def enable_pdf_drop(widget, variable: ctk.StringVar, on_change=None) -> bool:
    """Enable dropping a single PDF file onto an entry."""
    def handle(event):
        raw = _first_drop_path(widget, event)
        if not raw:
            return
        path = Path(raw)
        if path.is_dir():
            return
        if path.suffix.lower() != ".pdf":
            messagebox.showerror("Invalid file", "Please drop a PDF file.")
            return
        variable.set(str(path))
        if on_change:
            on_change()

    return _register_drop_target(widget, handle)


# Shared modern PaperKit palette
APP_BG = "#0B1120"
SURFACE = "#0F1B33"
SURFACE_ALT = "#162640"
SURFACE_SOFT = "#1E3356"
BORDER = "#2B4A70"
TEXT = "#F8FAFC"
MUTED = "#A8B5C7"
PRIMARY = "#2F6DF6"
PRIMARY_HOVER = "#1F55D8"
SECONDARY = "#1E3356"
SECONDARY_HOVER = "#25436C"
DANGER = "#E11D48"
DANGER_HOVER = "#BE123C"


def card(master, **kwargs) -> ctk.CTkFrame:
    kwargs.setdefault("fg_color", SURFACE)
    kwargs.setdefault("corner_radius", 16)
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("border_color", BORDER)
    return ctk.CTkFrame(master, **kwargs)


def subcard(master, **kwargs) -> ctk.CTkFrame:
    kwargs.setdefault("fg_color", SURFACE_ALT)
    kwargs.setdefault("corner_radius", 14)
    return ctk.CTkFrame(master, **kwargs)


def style_button(button: ctk.CTkButton, kind: str = "primary") -> None:
    if kind == "secondary":
        button.configure(fg_color=SECONDARY, hover_color=SECONDARY_HOVER)
    elif kind == "danger":
        button.configure(fg_color=DANGER, hover_color=DANGER_HOVER)
    else:
        button.configure(fg_color=PRIMARY, hover_color=PRIMARY_HOVER)


def apply_treeview_style() -> None:
    from tkinter import ttk
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(
        "Treeview",
        background="#0F1B33",
        foreground="#F8FAFC",
        fieldbackground="#0F1B33",
        bordercolor="#2B4A70",
        rowheight=26,
    )
    style.configure(
        "Treeview.Heading",
        background="#162640",
        foreground="#F8FAFC",
        bordercolor="#2B4A70",
        relief="flat",
    )
    style.map(
        "Treeview",
        background=[("selected", "#2F6DF6")],
        foreground=[("selected", "#FFFFFF")],
    )


def resource_path(relative_path: str) -> str:
    from pathlib import Path
    import sys

    if hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)  # PyInstaller _internal folder
    else:
        base_path = Path(__file__).resolve().parent.parent

    return str(base_path / relative_path)

def choose_folder(title: str) -> str:
    return filedialog.askdirectory(title=title) or ""


def choose_pdf(title: str) -> str:
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    ) or ""


def open_folder(path: str | Path) -> None:
    path = Path(path)
    if not path.exists():
        messagebox.showerror("Folder not found", f"This folder does not exist:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')


class LogMixin:
    log_box: ctk.CTkTextbox

    def log(self, message: str) -> None:
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def thread_safe_log(self, message: str) -> None:
        self.after(0, lambda: self.log(message))

    def clear_log(self) -> None:
        self.log_box.delete("1.0", "end")


def set_progress(progress_bar: ctk.CTkProgressBar, label: ctk.CTkLabel, current: int, total: int) -> None:
    if total <= 0:
        progress_bar.set(0)
        label.configure(text="0 / 0")
        return
    progress_bar.set(current / total)
    label.configure(text=f"{current} / {total}")


def count_pdfs(folder: str | Path) -> int:
    try:
        path = Path(folder)
        if not path.exists() or not path.is_dir():
            return 0
        return len(list(path.glob('*.pdf')))
    except Exception:
        return 0


def validate_input_folder(path: str | Path) -> tuple[bool, str]:
    path = Path(str(path).strip())
    if not str(path):
        return False, 'Not selected'
    if not path.exists():
        return False, 'Folder not found'
    if not path.is_dir():
        return False, 'Not a folder'
    n = count_pdfs(path)
    if n == 0:
        return False, 'No PDFs found'
    return True, f'{n} PDF(s) found'


def validate_output_folder(path: str | Path) -> tuple[bool, str]:
    path = Path(str(path).strip())
    if not str(path):
        return False, 'Not selected'
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / '.paperkit_write_test'
        probe.write_text('ok', encoding='utf-8')
        probe.unlink(missing_ok=True)
        return True, 'Output folder ready'
    except Exception:
        return False, 'Output folder not writable'


def set_badge(label: ctk.CTkLabel, ok: bool, text: str) -> None:
    label.configure(
        text=(('✓ ' if ok else '⚠ ') + text),
        text_color=('#7DD3FC' if ok else '#FCA5A5'),
    )


def make_summary_text(title: str, summary: dict[str, object]) -> str:
    parts = [title]
    for key, value in summary.items():
        parts.append(f'{key}: {value}')
    return '  •  '.join(parts)
