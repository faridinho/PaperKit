from __future__ import annotations

import threading
from pathlib import Path
from threading import Event
from tkinter import messagebox

import customtkinter as ctk

from pdf_tools import GHOSTSCRIPT_DOWNLOAD_URL, PDF_SETTINGS, compress_pdfs_in_folder, list_pdfs, export_operation_report
from tabs.common import LogMixin, choose_folder, open_folder, set_progress, card, validate_input_folder, validate_output_folder, set_badge, make_summary_text, enable_folder_drop, TEXT, MUTED, SECONDARY, SECONDARY_HOVER, DANGER, DANGER_HOVER, PRIMARY, PRIMARY_HOVER


class CompressTab(ctk.CTkFrame, LogMixin):
    def __init__(self, master) -> None:
        super().__init__(master, fg_color="transparent")
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "output" / "compressed"))
        self.quality_var = ctk.StringVar(value="ebook")
        self.gs_var = ctk.StringVar(value="")
        self.cancel_event = Event()
        self._build()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        folder = card(self)
        folder.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        folder.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folder, text="Input folder").grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.input_entry = ctk.CTkEntry(folder, textvariable=self.input_var, placeholder_text="Browse or drag a PDF folder here")
        self.input_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folder, text="Browse", command=self.choose_input, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=2, padx=12, pady=10)

        ctk.CTkLabel(folder, text="Output folder").grid(row=1, column=0, padx=12, pady=10, sticky="w")
        self.output_entry = ctk.CTkEntry(folder, textvariable=self.output_var, placeholder_text="Browse or drag an output folder here")
        self.output_entry.grid(row=1, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folder, text="Browse", command=self.choose_output, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=1, column=2, padx=12, pady=10)

        self.input_badge = ctk.CTkLabel(folder, text="⚠ Input not selected", text_color="#FCA5A5", anchor="w")
        self.input_badge.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")
        self.output_badge = ctk.CTkLabel(folder, text="⚠ Output not checked", text_color="#FCA5A5", anchor="w")
        self.output_badge.grid(row=2, column=2, padx=12, pady=(0, 8), sticky="e")

        options = card(self)
        options.grid(row=1, column=0, padx=10, pady=6, sticky="ew")
        options.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(options, text="Compression quality").grid(row=0, column=0, padx=12, pady=10, sticky="w")
        ctk.CTkOptionMenu(options, variable=self.quality_var, values=list(PDF_SETTINGS.keys())).grid(row=0, column=1, padx=8, pady=10, sticky="w")

        ctk.CTkLabel(options, text="Ghostscript executable").grid(row=1, column=0, padx=12, pady=10, sticky="w")
        ctk.CTkEntry(options, textvariable=self.gs_var, placeholder_text="Optional: gswin64c, gs, or full path").grid(row=1, column=1, padx=8, pady=10, sticky="ew")

        help_text = (
            "Leave Ghostscript empty if gswin64c/gs is already in PATH. "
            f"If compression fails, install Ghostscript from: {GHOSTSCRIPT_DOWNLOAD_URL}"
        )
        ctk.CTkLabel(options, text=help_text, wraplength=820, justify="left").grid(row=2, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="w")

        actions = card(self)
        actions.grid(row=2, column=0, padx=10, pady=6, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        self.start_button = ctk.CTkButton(actions, text="Start compression", command=self.start, fg_color=PRIMARY, hover_color=PRIMARY_HOVER)
        self.start_button.grid(row=0, column=1, padx=8, pady=10)
        self.cancel_button = ctk.CTkButton(actions, text="Cancel", command=self.cancel, state="disabled", fg_color=DANGER, hover_color=DANGER_HOVER)
        self.cancel_button.grid(row=0, column=2, padx=8, pady=10)
        ctk.CTkButton(actions, text="Open output folder", command=lambda: open_folder(self.output_var.get()), fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=3, padx=8, pady=10)
        ctk.CTkButton(actions, text="Clear log", command=self.clear_log, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=4, padx=8, pady=10)

        progress_frame = card(self)
        progress_frame.grid(row=3, column=0, padx=10, pady=6, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(progress_frame)
        self.progress.grid(row=0, column=0, padx=12, pady=10, sticky="ew")
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(progress_frame, text="0 / 0", width=80)
        self.progress_label.grid(row=0, column=1, padx=12, pady=10)

        self.summary_card = card(self)
        self.summary_card.grid(row=4, column=0, padx=10, pady=(6, 0), sticky="ew")
        self.summary_card.grid_columnconfigure(0, weight=1)
        self.summary_label = ctk.CTkLabel(self.summary_card, text="Summary will appear here after compression finishes.", text_color=MUTED, anchor="w", justify="left")
        self.summary_label.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        self.log_box = ctk.CTkTextbox(self, wrap="word", fg_color="#0A1020", border_width=1, border_color="#2B4A70", text_color=TEXT)
        self.log_box.grid(row=5, column=0, padx=10, pady=10, sticky="nsew")
        self.log("Choose folders and click Start compression.")
        self._update_badges()
        self._enable_drag_and_drop()

    def _enable_drag_and_drop(self) -> None:
        enable_folder_drop(self.input_entry, self.input_var, self._update_badges)
        enable_folder_drop(self.output_entry, self.output_var, self._update_badges)


    def _update_badges(self) -> None:
        ok, text = validate_input_folder(self.input_var.get())
        set_badge(self.input_badge, ok, text)
        ok, text = validate_output_folder(self.output_var.get())
        set_badge(self.output_badge, ok, text)

    def _show_summary(self, title: str, summary: dict[str, object]) -> None:
        self.summary_label.configure(text=make_summary_text(title, summary), text_color=TEXT)

    def choose_input(self) -> None:
        folder = choose_folder("Choose input folder")
        if folder:
            self.input_var.set(folder)
            self._update_badges()

    def choose_output(self) -> None:
        folder = choose_folder("Choose output folder")
        if folder:
            self.output_var.set(folder)
            self._update_badges()

    def _validate(self) -> bool:
        self._update_badges()
        if not self.input_var.get().strip():
            messagebox.showerror("Missing input", "Please choose an input folder.")
            return False
        if not self.output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return False
        return True

    def _update_progress(self, current: int, total: int) -> None:
        self.after(0, lambda: set_progress(self.progress, self.progress_label, current, total))

    def start(self) -> None:
        if not self._validate():
            return
        self.clear_log()
        self.cancel_event.clear()
        self.start_button.configure(state="disabled", text="Compressing...")
        self.cancel_button.configure(state="normal")
        threading.Thread(target=self._worker, daemon=True).start()

    def cancel(self) -> None:
        self.cancel_event.set()
        self.log("Cancel requested. Current Ghostscript process will be stopped if possible.")
        self.cancel_button.configure(state="disabled")

    def _worker(self) -> None:
        try:
            input_dir = Path(self.input_var.get())
            output_dir = Path(self.output_var.get())
            total_files = len(list_pdfs(input_dir))
            compressed = compress_pdfs_in_folder(
                input_dir=input_dir,
                output_dir=output_dir,
                quality=self.quality_var.get(),
                gs_executable=self.gs_var.get().strip() or None,
                log=self.thread_safe_log,
                progress=self._update_progress,
                cancel_event=self.cancel_event,
            )
            report = output_dir / "paperkit_compression_report.xlsx"
            export_operation_report(
                report,
                "PaperKit Compression Report",
                {
                    "Input PDFs": total_files,
                    "Compressed": len(compressed),
                    "Cancelled": "Yes" if self.cancel_event.is_set() else "No",
                    "Quality": self.quality_var.get(),
                    "Output folder": str(output_dir),
                },
                [{"Compressed file": str(p)} for p in compressed],
                log=self.thread_safe_log,
            )
            self.after(0, lambda: self._show_summary("Compression finished" if not self.cancel_event.is_set() else "Compression cancelled", {"Input PDFs": total_files, "Compressed": len(compressed), "Report": report.name}))
            if self.cancel_event.is_set():
                self.after(0, lambda: messagebox.showinfo("Cancelled", f"Compression was cancelled.\n\nReport created:\n{report}"))
            else:
                self.after(0, lambda: messagebox.showinfo("Done", f"Compression finished.\n\nReport created:\n{report}"))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self.start_button.configure(state="normal", text="Start compression"))
            self.after(0, lambda: self.cancel_button.configure(state="disabled"))
