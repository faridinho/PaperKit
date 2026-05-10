from __future__ import annotations

import threading
from pathlib import Path
from tkinter import messagebox, ttk

import customtkinter as ctk

from pdf_tools import (
    apply_rename_plan,
    build_rename_plan,
    export_duplicate_report,
    export_operation_report,
    find_duplicate_pdfs_by_hash,
)
from tabs.common import LogMixin, choose_folder, open_folder, set_progress, card, subcard, apply_treeview_style, validate_input_folder, validate_output_folder, set_badge, make_summary_text, enable_folder_drop, TEXT, MUTED


class RenameTab(ctk.CTkFrame, LogMixin):
    def __init__(self, master) -> None:
        super().__init__(master, fg_color="transparent")

        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.in_place_var = ctk.BooleanVar(value=False)
        self.add_numbering_var = ctk.BooleanVar(value=True)
        self.start_var = ctk.StringVar(value="1")
        self.lowercase_rest_var = ctk.BooleanVar(value=False)
        self.rename_plan: list[dict] = []

        # Used to refresh the preview automatically after option changes.
        self.preview_after_id: str | None = None
        self.preview_running = False

        self._build()
        self._connect_option_traces()
        self._update_numbering_state()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        folder_frame = card(self)
        folder_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        folder_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folder_frame, text="Input folder").grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.input_entry = ctk.CTkEntry(folder_frame, textvariable=self.input_var, placeholder_text="Browse or drag a PDF folder here")
        self.input_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folder_frame, text="Browse", command=self.choose_input).grid(row=0, column=2, padx=12, pady=10)

        ctk.CTkLabel(folder_frame, text="Output folder").grid(row=1, column=0, padx=12, pady=10, sticky="w")
        self.output_entry = ctk.CTkEntry(folder_frame, textvariable=self.output_var, placeholder_text="Browse or drag an output folder here")
        self.output_entry.grid(row=1, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(folder_frame, text="Browse", command=self.choose_output).grid(row=1, column=2, padx=12, pady=10)

        self.input_badge = ctk.CTkLabel(folder_frame, text="⚠ Input not selected", text_color="#FCA5A5", anchor="w")
        self.input_badge.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")
        self.output_badge = ctk.CTkLabel(folder_frame, text="⚠ Output not checked", text_color="#FCA5A5", anchor="w")
        self.output_badge.grid(row=2, column=2, padx=12, pady=(0, 8), sticky="e")

        options = card(self)
        options.grid(row=1, column=0, padx=10, pady=6, sticky="ew")
        options.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.in_place_check = ctk.CTkCheckBox(
            options,
            text="Rename originals directly",
            variable=self.in_place_var,
            command=self._option_changed,
        )
        self.in_place_check.grid(row=0, column=0, padx=12, pady=10, sticky="w")

        self.add_numbering_check = ctk.CTkCheckBox(
            options,
            text="Add numbering",
            variable=self.add_numbering_var,
            command=self._option_changed,
        )
        self.add_numbering_check.grid(row=0, column=1, padx=12, pady=10, sticky="w")

        ctk.CTkLabel(options, text="Start number").grid(row=0, column=2, padx=(12, 4), pady=10, sticky="e")
        self.start_entry = ctk.CTkEntry(options, textvariable=self.start_var, width=90)
        self.start_entry.grid(row=0, column=3, padx=(4, 12), pady=10, sticky="w")

        self.lowercase_check = ctk.CTkCheckBox(
            options,
            text="Old title style: lowercase rest",
            variable=self.lowercase_rest_var,
            command=self._option_changed,
        )
        self.lowercase_check.grid(row=1, column=0, columnspan=2, padx=12, pady=10, sticky="w")

        action = card(self)
        action.grid(row=2, column=0, padx=10, pady=6, sticky="ew")
        action.grid_columnconfigure(0, weight=1)

        self.preview_button = ctk.CTkButton(action, text="Preview renaming", command=self.start_preview)
        self.preview_button.grid(row=0, column=1, padx=8, pady=10)
        self.apply_button = ctk.CTkButton(action, text="Apply renaming", command=self.start_apply)
        self.apply_button.grid(row=0, column=2, padx=8, pady=10)
        self.duplicates_button = ctk.CTkButton(action, text="Detect duplicates", command=self.start_duplicate_scan)
        self.duplicates_button.grid(row=0, column=3, padx=8, pady=10)
        ctk.CTkButton(action, text="Open output folder", command=lambda: open_folder(self.output_var.get())).grid(row=0, column=4, padx=8, pady=10)

        progress_frame = card(self)
        progress_frame.grid(row=3, column=0, padx=10, pady=6, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(progress_frame)
        self.progress.grid(row=0, column=0, padx=12, pady=10, sticky="ew")
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(progress_frame, text="0 / 0", width=80)
        self.progress_label.grid(row=0, column=1, padx=12, pady=10)

        table_frame = card(self)
        table_frame.grid(row=4, column=0, padx=10, pady=6, sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        apply_treeview_style()
        columns = ("old", "title", "new", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=8)
        self.tree.heading("old", text="Old filename")
        self.tree.heading("title", text="Extracted title")
        self.tree.heading("new", text="New filename")
        self.tree.heading("status", text="Status")
        self.tree.column("old", width=180)
        self.tree.column("title", width=300)
        self.tree.column("new", width=300)
        self.tree.column("status", width=180)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.summary_card = card(self)
        self.summary_card.grid(row=5, column=0, padx=10, pady=(6, 0), sticky="ew")
        self.summary_card.grid_columnconfigure(0, weight=1)
        self.summary_label = ctk.CTkLabel(self.summary_card, text="Summary will appear here after an operation finishes.", text_color=MUTED, anchor="w", justify="left")
        self.summary_label.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        self.log_box = ctk.CTkTextbox(self, height=105, wrap="word", fg_color="#0A1020", border_width=1, border_color="#2B4A70", text_color=TEXT)
        self.log_box.grid(row=6, column=0, padx=10, pady=10, sticky="ew")
        self.log("Preview first, then apply the renaming after checking the table.")
        self._update_badges()
        self._enable_drag_and_drop()

    def _enable_drag_and_drop(self) -> None:
        enable_folder_drop(self.input_entry, self.input_var, self._on_input_dropped)
        enable_folder_drop(self.output_entry, self.output_var, self._on_output_dropped)

    def _on_input_dropped(self) -> None:
        self.rename_plan = []
        self._clear_table()
        self._update_badges()
        self.log("Input folder dropped. Run Preview renaming again.")

    def _on_output_dropped(self) -> None:
        self._update_badges()
        if self.rename_plan:
            self._option_changed()

    def _connect_option_traces(self) -> None:
        # This catches manual typing in the Start number field.
        self.start_var.trace_add("write", lambda *_: self._option_changed())

    def _option_changed(self) -> None:
        """Refresh the preview when rename options change."""
        self._update_numbering_state()

        # Only auto-refresh if a preview already exists. Before that, the user can
        # choose options freely and click Preview once.
        if not self.rename_plan:
            return

        # Debounce rapid typing in the Start number field.
        if self.preview_after_id is not None:
            self.after_cancel(self.preview_after_id)
        self.preview_after_id = self.after(450, self._refresh_preview_after_option_change)

    def _update_numbering_state(self) -> None:
        if self.add_numbering_var.get():
            self.start_entry.configure(state="normal")
        else:
            self.start_entry.configure(state="disabled")

    def _refresh_preview_after_option_change(self) -> None:
        self.preview_after_id = None
        if self.preview_running:
            return
        self.start_preview(auto=True)


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
            self.rename_plan = []
            self._clear_table()
            self.log("Input folder changed. Run Preview renaming again.")
            self._update_badges()

    def choose_output(self) -> None:
        folder = choose_folder("Choose output folder")
        if folder:
            self.output_var.set(folder)
            self._update_badges()
            if self.rename_plan:
                self._option_changed()

    def _get_start_number(self) -> int:
        if not self.add_numbering_var.get():
            return 1
        try:
            return int(self.start_var.get())
        except ValueError:
            raise ValueError("Start number must be an integer.")

    def _validate(self) -> bool:
        self._update_badges()
        if not self.input_var.get().strip():
            messagebox.showerror("Missing input", "Please choose an input folder.")
            return False
        if not self.output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return False
        return True

    def _set_buttons(self, state: str) -> None:
        self.preview_button.configure(state=state)
        self.apply_button.configure(state=state)
        self.duplicates_button.configure(state=state)

    def _update_progress(self, current: int, total: int) -> None:
        self.after(0, lambda: set_progress(self.progress, self.progress_label, current, total))

    def _clear_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _fill_table(self) -> None:
        self._clear_table()
        for row in self.rename_plan:
            self.tree.insert("", "end", values=(row.get("old_filename"), row.get("title"), row.get("new_filename"), row.get("status")))

    def start_preview(self, auto: bool = False) -> None:
        if not self._validate():
            return
        try:
            start_number = self._get_start_number()
        except ValueError as exc:
            # While the user is typing, avoid noisy popups on automatic refresh.
            if not auto:
                messagebox.showerror("Invalid start number", str(exc))
            return

        if self.preview_running:
            return

        self.preview_running = True
        if not auto:
            self.clear_log()
        else:
            self.log("Options changed. Refreshing preview...")
        self._set_buttons("disabled")
        threading.Thread(target=self._preview_worker, args=(start_number,), daemon=True).start()

    def _preview_worker(self, start_number: int) -> None:
        try:
            self.rename_plan = build_rename_plan(
                input_dir=Path(self.input_var.get()),
                output_dir=Path(self.output_var.get()),
                start_number=start_number,
                add_numbering=self.add_numbering_var.get(),
                in_place=self.in_place_var.get(),
                lowercase_rest=self.lowercase_rest_var.get(),
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            self.after(0, self._fill_table)
            ready = sum(1 for row in self.rename_plan if not str(row.get("status", "")).startswith("ERROR"))
            self.after(0, lambda: self._show_summary("Preview complete", {"Files scanned": len(self.rename_plan), "Ready": ready, "Needs review": len(self.rename_plan) - ready}))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.preview_running = False
            self.after(0, lambda: self._set_buttons("normal"))

    def start_apply(self) -> None:
        if not self.rename_plan:
            messagebox.showerror("No preview", "Please run Preview renaming first.")
            return
        confirm = messagebox.askyesno("Apply renaming", "Apply the rename plan shown in the preview table?")
        if not confirm:
            return
        self.clear_log()
        self._set_buttons("disabled")
        threading.Thread(target=self._apply_worker, daemon=True).start()

    def _apply_worker(self) -> None:
        try:
            renamed = apply_rename_plan(
                self.rename_plan,
                in_place=self.in_place_var.get(),
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            report = Path(self.output_var.get()) / "paperkit_rename_report.xlsx"
            export_operation_report(
                report,
                "PaperKit Rename Report",
                {
                    "Files in preview": len(self.rename_plan),
                    "Files renamed/copied": len(renamed),
                    "Mode": "In-place rename" if self.in_place_var.get() else "Safe copy rename",
                    "Output folder": self.output_var.get(),
                },
                self.rename_plan,
                log=self.thread_safe_log,
            )
            self.after(0, lambda: self._show_summary("Renaming finished", {"Previewed": len(self.rename_plan), "Renamed/copied": len(renamed), "Report": report.name}))
            self.after(0, lambda: messagebox.showinfo("Done", f"Renaming finished.\n\nReport created:\n{report}"))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self._set_buttons("normal"))

    def start_duplicate_scan(self) -> None:
        if not self._validate():
            return
        self.clear_log()
        self._set_buttons("disabled")
        threading.Thread(target=self._duplicate_worker, daemon=True).start()

    def _duplicate_worker(self) -> None:
        try:
            rows = find_duplicate_pdfs_by_hash(
                Path(self.input_var.get()),
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            if rows:
                out = Path(self.output_var.get()) / "duplicate_report.xlsx"
                export_duplicate_report(rows, out, log=self.thread_safe_log)
                groups = len(set(row.get("Group") for row in rows))
                self.after(0, lambda: self._show_summary("Duplicate scan finished", {"Duplicate groups": groups, "Duplicate files": len(rows), "Report": out.name}))
                self.after(0, lambda: messagebox.showinfo("Done", f"Duplicate report created:\n{out}"))
            else:
                self.after(0, lambda: self._show_summary("Duplicate scan finished", {"Duplicate groups": 0, "Duplicate files": 0}))
                self.after(0, lambda: messagebox.showinfo("Done", "No exact duplicate PDFs found."))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self._set_buttons("normal"))
