from __future__ import annotations

import threading
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from pdf_tools import DEFAULT_METADATA_FIELDS, METADATA_FIELDS, export_metadata_to_excel, list_pdfs, export_operation_report
from tabs.common import LogMixin, choose_folder, open_folder, set_progress, card, validate_input_folder, validate_output_folder, set_badge, make_summary_text, enable_folder_drop, TEXT, MUTED, SECONDARY, SECONDARY_HOVER, PRIMARY, PRIMARY_HOVER


class MetadataTab(ctk.CTkFrame, LogMixin):
    def __init__(self, master) -> None:
        super().__init__(master, fg_color="transparent")
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.field_vars: dict[str, ctk.BooleanVar] = {}
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

        fields = card(self)
        fields.grid(row=1, column=0, padx=10, pady=6, sticky="ew")
        fields.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ctk.CTkLabel(fields, text="Choose metadata fields", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        for i, field in enumerate(METADATA_FIELDS):
            var = ctk.BooleanVar(value=field in DEFAULT_METADATA_FIELDS)
            self.field_vars[field] = var
            row = 1 + i // 4
            col = i % 4
            ctk.CTkCheckBox(fields, text=field, variable=var).grid(row=row, column=col, padx=12, pady=6, sticky="w")

        shortcuts = card(self)
        shortcuts.grid(row=2, column=0, padx=10, pady=6, sticky="ew")
        shortcuts.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(shortcuts, text="Default fields", fg_color=SECONDARY, hover_color=SECONDARY_HOVER, command=self.default_fields).grid(row=0, column=1, padx=8, pady=10)
        ctk.CTkButton(shortcuts, text="Select all", fg_color=SECONDARY, hover_color=SECONDARY_HOVER, command=self.select_all).grid(row=0, column=2, padx=8, pady=10)
        ctk.CTkButton(shortcuts, text="Clear all", fg_color=SECONDARY, hover_color=SECONDARY_HOVER, command=self.clear_fields).grid(row=0, column=3, padx=8, pady=10)
        self.extract_button = ctk.CTkButton(shortcuts, text="Export metadata to Excel", command=self.start, fg_color=PRIMARY, hover_color=PRIMARY_HOVER)
        self.extract_button.grid(row=0, column=4, padx=8, pady=10)
        ctk.CTkButton(shortcuts, text="Open output folder", fg_color=SECONDARY, hover_color=SECONDARY_HOVER, command=lambda: open_folder(self.output_var.get())).grid(row=0, column=5, padx=8, pady=10)

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
        self.summary_label = ctk.CTkLabel(self.summary_card, text="Summary will appear here after metadata export finishes.", text_color=MUTED, anchor="w", justify="left")
        self.summary_label.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        self.log_box = ctk.CTkTextbox(self, wrap="word", fg_color="#0A1020", border_width=1, border_color="#2B4A70", text_color=TEXT)
        self.log_box.grid(row=5, column=0, padx=10, pady=10, sticky="nsew")
        self.log("Select the metadata fields you want, then export to Excel.")
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

    def selected_fields(self) -> list[str]:
        return [field for field, var in self.field_vars.items() if var.get()]

    def default_fields(self) -> None:
        for field, var in self.field_vars.items():
            var.set(field in DEFAULT_METADATA_FIELDS)

    def select_all(self) -> None:
        for var in self.field_vars.values():
            var.set(True)

    def clear_fields(self) -> None:
        for var in self.field_vars.values():
            var.set(False)

    def _validate(self) -> bool:
        self._update_badges()
        if not self.input_var.get().strip():
            messagebox.showerror("Missing input", "Please choose an input folder.")
            return False
        if not self.output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return False
        if not self.selected_fields():
            messagebox.showerror("No fields selected", "Please select at least one metadata field.")
            return False
        return True

    def _update_progress(self, current: int, total: int) -> None:
        self.after(0, lambda: set_progress(self.progress, self.progress_label, current, total))

    def start(self) -> None:
        if not self._validate():
            return
        self.clear_log()
        self.extract_button.configure(state="disabled", text="Exporting...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self) -> None:
        try:
            out = Path(self.output_var.get()) / "pdf_metadata.xlsx"
            input_dir = Path(self.input_var.get())
            fields = self.selected_fields()
            export_metadata_to_excel(
                input_dir=input_dir,
                output_file=out,
                fields=fields,
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            report = Path(self.output_var.get()) / "paperkit_metadata_report.xlsx"
            export_operation_report(
                report,
                "PaperKit Metadata Export Report",
                {
                    "Input PDFs": len(list_pdfs(input_dir)),
                    "Selected fields": ", ".join(fields),
                    "Excel file": str(out),
                },
                log=self.thread_safe_log,
            )
            self.after(0, lambda: self._show_summary("Metadata exported", {"PDFs": len(list_pdfs(input_dir)), "Fields": len(fields), "Report": report.name}))
            self.after(0, lambda: messagebox.showinfo("Done", f"Metadata exported:\n{out}\n\nReport created:\n{report}"))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self.extract_button.configure(state="normal", text="Export metadata to Excel"))
