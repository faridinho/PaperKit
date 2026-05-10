from __future__ import annotations

import threading
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from pdf_tools import export_citations, list_pdfs, export_operation_report
from tabs.common import (
    LogMixin,
    choose_folder,
    open_folder,
    set_progress,
    card,
    TEXT,
    MUTED,
    PRIMARY,
    PRIMARY_HOVER,
    SECONDARY,
    SECONDARY_HOVER,
    validate_input_folder,
    validate_output_folder,
    set_badge,
    make_summary_text,
    enable_folder_drop,
    DANGER,
)


class CitationTab(ctk.CTkFrame, LogMixin):
    def __init__(self, master) -> None:
        super().__init__(master, fg_color="transparent")
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.ris_var = ctk.BooleanVar(value=True)
        self.bibtex_var = ctk.BooleanVar(value=True)
        self._build()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        warning = card(self)
        warning.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        warning.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            warning,
            text="Citation export uses PDF metadata",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=TEXT,
            anchor="w",
        ).grid(row=0, column=0, padx=14, pady=(14, 4), sticky="ew")

        ctk.CTkLabel(
            warning,
            text=(
                "PaperKit automatically creates RIS and/or BibTeX files from available PDF metadata "
                "and detectable first-page information such as title, author, year, and DOI. "
                "Many PDFs contain incomplete or incorrect metadata, so always review imported records "
                "in Mendeley, EndNote, Zotero, JabRef, or your reference manager."
            ),
            wraplength=860,
            justify="left",
            text_color=MUTED,
            anchor="w",
        ).grid(row=1, column=0, padx=14, pady=(0, 14), sticky="ew")

        folders = card(self)
        folders.grid(row=1, column=0, padx=10, pady=8, sticky="ew")
        folders.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folders, text="Input folder", text_color=TEXT).grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.input_entry = ctk.CTkEntry(folders, textvariable=self.input_var, placeholder_text="Browse or drag a PDF folder here")
        self.input_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folders, text="Browse", command=self.choose_input, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=2, padx=12, pady=10)

        ctk.CTkLabel(folders, text="Output folder", text_color=TEXT).grid(row=1, column=0, padx=12, pady=10, sticky="w")
        self.output_entry = ctk.CTkEntry(folders, textvariable=self.output_var, placeholder_text="Browse or drag an output folder here")
        self.output_entry.grid(row=1, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folders, text="Browse", command=self.choose_output, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=1, column=2, padx=12, pady=10)

        self.input_badge = ctk.CTkLabel(folders, text="⚠ Input not selected", text_color="#FCA5A5", anchor="w")
        self.input_badge.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")
        self.output_badge = ctk.CTkLabel(folders, text="⚠ Output not checked", text_color="#FCA5A5", anchor="w")
        self.output_badge.grid(row=2, column=2, padx=12, pady=(0, 8), sticky="e")

        options = card(self)
        options.grid(row=2, column=0, padx=10, pady=8, sticky="ew")
        options.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(options, text="Export formats", font=ctk.CTkFont(weight="bold"), text_color=TEXT).grid(row=0, column=0, padx=12, pady=12, sticky="w")
        ctk.CTkCheckBox(options, text="RIS for Mendeley, EndNote, Zotero", variable=self.ris_var).grid(row=0, column=1, padx=12, pady=12, sticky="w")
        ctk.CTkCheckBox(options, text="BibTeX for LaTeX, JabRef, Zotero", variable=self.bibtex_var).grid(row=0, column=2, padx=12, pady=12, sticky="w")

        actions = card(self)
        actions.grid(row=3, column=0, padx=10, pady=8, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        self.export_button = ctk.CTkButton(
            actions,
            text="Export citations",
            command=self.start,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
        )
        self.export_button.grid(row=0, column=1, padx=8, pady=10)
        ctk.CTkButton(
            actions,
            text="Open output folder",
            command=lambda: open_folder(self.output_var.get()),
            fg_color=SECONDARY,
            hover_color=SECONDARY_HOVER,
        ).grid(row=0, column=2, padx=8, pady=10)

        self.progress = ctk.CTkProgressBar(actions)
        self.progress.grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 12), sticky="ew")
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(actions, text="0 / 0", text_color=TEXT, width=70)
        self.progress_label.grid(row=1, column=2, padx=12, pady=(0, 12), sticky="e")

        self.summary_card = card(self)
        self.summary_card.grid(row=4, column=0, padx=10, pady=(6, 0), sticky="ew")
        self.summary_card.grid_columnconfigure(0, weight=1)
        self.summary_label = ctk.CTkLabel(self.summary_card, text="Summary will appear here after citation export finishes.", text_color=MUTED, anchor="w", justify="left")
        self.summary_label.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        self.log_box = ctk.CTkTextbox(self, wrap="word", fg_color="#0A1020", border_width=1, border_color="#2B4A70", text_color=TEXT)
        self.log_box.grid(row=5, column=0, padx=10, pady=10, sticky="nsew")
        self.log("Choose a folder and export citations. RIS and BibTeX are generated automatically from available PDF metadata.")
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
        if not self.ris_var.get() and not self.bibtex_var.get():
            messagebox.showerror("No format selected", "Please select RIS, BibTeX, or both.")
            return False
        return True

    def _update_progress(self, current: int, total: int) -> None:
        self.after(0, lambda: set_progress(self.progress, self.progress_label, current, total))

    def start(self) -> None:
        if not self._validate():
            return
        self.clear_log()
        self.export_button.configure(state="disabled", text="Exporting...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self) -> None:
        try:
            input_dir = Path(self.input_var.get())
            output_dir = Path(self.output_var.get())
            outputs = export_citations(
                input_dir=input_dir,
                output_dir=output_dir,
                export_ris=self.ris_var.get(),
                export_bibtex=self.bibtex_var.get(),
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            report = output_dir / "paperkit_citation_report.xlsx"
            export_operation_report(
                report,
                "PaperKit Citation Export Report",
                {
                    "Input PDFs": len(list_pdfs(input_dir)),
                    "RIS exported": "Yes" if self.ris_var.get() else "No",
                    "BibTeX exported": "Yes" if self.bibtex_var.get() else "No",
                    "Output files": ", ".join(str(p.name) for p in outputs),
                },
                log=self.thread_safe_log,
            )
            self.after(0, lambda: self._show_summary("Citations exported", {"PDFs": len(list_pdfs(input_dir)), "Files": len(outputs), "Report": report.name}))
            msg = "Citation export finished."
            if outputs:
                msg += "\n\n" + "\n".join(str(p) for p in outputs)
            msg += f"\n\nReport created:\n{report}"
            self.after(0, lambda: messagebox.showinfo("Done", msg))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self.export_button.configure(state="normal", text="Export citations"))
