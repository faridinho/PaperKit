from __future__ import annotations

import threading
from io import BytesIO
from pathlib import Path
from tkinter import messagebox, ttk

import customtkinter as ctk
from PIL import Image

from pdf_tools import extract_selected_images_from_pdf, get_image_bytes, scan_embedded_images, export_operation_report
from tabs.common import LogMixin, choose_folder, choose_pdf, open_folder, set_progress, card, apply_treeview_style, validate_output_folder, set_badge, make_summary_text, enable_folder_drop, enable_pdf_drop, TEXT, MUTED, SECONDARY, SECONDARY_HOVER, PRIMARY, PRIMARY_HOVER


class ImageTab(ctk.CTkFrame, LogMixin):
    def __init__(self, master) -> None:
        super().__init__(master, fg_color="transparent")
        self.pdf_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "extracted_images"))
        self.images: list[dict] = []
        self.preview_image = None
        self._build()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        files = card(self)
        files.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        files.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(files, text="Single PDF file", text_color=TEXT).grid(row=0, column=0, padx=12, pady=10, sticky="w")
        self.pdf_entry = ctk.CTkEntry(files, textvariable=self.pdf_var, placeholder_text="Browse or drag a single PDF here")
        self.pdf_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(files, text="Browse", command=self.choose_pdf_file, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=2, padx=12, pady=10)

        ctk.CTkLabel(files, text="Output folder", text_color=TEXT).grid(row=1, column=0, padx=12, pady=10, sticky="w")
        self.output_entry = ctk.CTkEntry(files, textvariable=self.output_var, placeholder_text="Browse or drag an output folder here")
        self.output_entry.grid(row=1, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(files, text="Browse", command=self.choose_output, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=1, column=2, padx=12, pady=10)

        self.file_badge = ctk.CTkLabel(files, text="⚠ PDF not selected", text_color="#FCA5A5", anchor="w")
        self.file_badge.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")
        self.output_badge = ctk.CTkLabel(files, text="⚠ Output not checked", text_color="#FCA5A5", anchor="w")
        self.output_badge.grid(row=2, column=2, padx=12, pady=(0, 8), sticky="e")

        actions = card(self)
        actions.grid(row=1, column=0, padx=10, pady=6, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        self.scan_button = ctk.CTkButton(actions, text="Scan images", command=self.start_scan, fg_color=PRIMARY, hover_color=PRIMARY_HOVER)
        self.scan_button.grid(row=0, column=1, padx=8, pady=10)
        ctk.CTkButton(actions, text="Select all", command=self.select_all, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=2, padx=8, pady=10)
        ctk.CTkButton(actions, text="Clear selection", command=self.clear_selection, fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=3, padx=8, pady=10)
        self.extract_button = ctk.CTkButton(actions, text="Extract selected images", command=self.start_extract, fg_color=PRIMARY, hover_color=PRIMARY_HOVER)
        self.extract_button.grid(row=0, column=4, padx=8, pady=10)
        ctk.CTkButton(actions, text="Open output folder", command=lambda: open_folder(self.output_var.get()), fg_color=SECONDARY, hover_color=SECONDARY_HOVER).grid(row=0, column=5, padx=8, pady=10)

        progress_frame = card(self)
        progress_frame.grid(row=2, column=0, padx=10, pady=6, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(progress_frame)
        self.progress.grid(row=0, column=0, padx=12, pady=10, sticky="ew")
        self.progress.set(0)
        self.progress_label = ctk.CTkLabel(progress_frame, text="0 / 0", width=80, text_color=TEXT)
        self.progress_label.grid(row=0, column=1, padx=12, pady=10)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=3, column=0, padx=10, pady=6, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=3)
        body.grid_rowconfigure(0, weight=1)

        table_frame = card(body)
        table_frame.grid(row=0, column=0, padx=(0, 8), pady=0, sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(table_frame, text="Images", font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT).grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        apply_treeview_style()
        columns = ("page", "num", "size")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        for key, text, width in [
            ("page", "Page", 60),
            ("num", "Image", 70),
            ("size", "Size", 110),
        ]:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, minwidth=50, stretch=True)
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(0, 12))
        self.tree.bind("<<TreeviewSelect>>", self.show_selected_preview)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(0, 12), padx=(0, 12))
        self.tree.configure(yscrollcommand=scrollbar.set)

        preview_frame = card(body)
        preview_frame.grid(row=0, column=1, padx=(8, 0), pady=0, sticky="nsew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(preview_frame, text="Preview", font=ctk.CTkFont(size=15, weight="bold"), text_color=TEXT).grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")
        self.preview_label = ctk.CTkLabel(
            preview_frame,
            text="Select an image",
            text_color=MUTED,
            fg_color="#0A1020",
            corner_radius=14,
        )
        self.preview_label.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

        self.summary_card = card(self)
        self.summary_card.grid(row=4, column=0, padx=10, pady=(6, 0), sticky="ew")
        self.summary_card.grid_columnconfigure(0, weight=1)
        self.summary_label = ctk.CTkLabel(self.summary_card, text="Summary will appear here after image scan or extraction.", text_color=MUTED, anchor="w", justify="left")
        self.summary_label.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        self.log_box = ctk.CTkTextbox(self, height=95, wrap="word", fg_color="#0A1020", border_width=1, border_color="#2B4A70", text_color=TEXT)
        self.log_box.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
        self.log("Scan a PDF first, then select images from the table and extract them.")
        self._update_badges()
        self._enable_drag_and_drop()

    def _enable_drag_and_drop(self) -> None:
        enable_pdf_drop(self.pdf_entry, self.pdf_var, self._on_pdf_dropped)
        enable_folder_drop(self.output_entry, self.output_var, self._update_badges)

    def _on_pdf_dropped(self) -> None:
        self._update_badges()
        self.images = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.preview_label.configure(image=None, text="Scan images to preview embedded images")
        self.preview_image = None
        self.log("PDF dropped. Click Scan images.")


    def _update_badges(self) -> None:
        pdf = Path(self.pdf_var.get().strip()) if self.pdf_var.get().strip() else None
        if pdf and pdf.exists() and pdf.suffix.lower() == ".pdf":
            set_badge(self.file_badge, True, "PDF selected")
        else:
            set_badge(self.file_badge, False, "PDF not selected")
        ok, text = validate_output_folder(self.output_var.get())
        set_badge(self.output_badge, ok, text)

    def _show_summary(self, title: str, summary: dict[str, object]) -> None:
        self.summary_label.configure(text=make_summary_text(title, summary), text_color=TEXT)

    def choose_pdf_file(self) -> None:
        path = choose_pdf("Choose one PDF file")
        if path:
            self.pdf_var.set(path)
            self._update_badges()

    def choose_output(self) -> None:
        folder = choose_folder("Choose output folder")
        if folder:
            self.output_var.set(folder)
            self._update_badges()

    def _validate_pdf(self) -> bool:
        self._update_badges()
        if not self.pdf_var.get().strip():
            messagebox.showerror("Missing PDF", "Please choose a single PDF file.")
            return False
        if not self.output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return False
        return True

    def _update_progress(self, current: int, total: int) -> None:
        self.after(0, lambda: set_progress(self.progress, self.progress_label, current, total))

    def _set_buttons(self, state: str) -> None:
        self.scan_button.configure(state=state)
        self.extract_button.configure(state=state)

    def start_scan(self) -> None:
        if not self._validate_pdf():
            return
        self.clear_log()
        self._set_buttons("disabled")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self) -> None:
        try:
            self.images = scan_embedded_images(Path(self.pdf_var.get()), log=self.thread_safe_log)
            self.after(0, self._fill_table)
            self.after(0, lambda: set_progress(self.progress, self.progress_label, len(self.images), len(self.images)))
            self.after(0, lambda: self._show_summary("Image scan finished", {"Images found": len(self.images), "PDF": Path(self.pdf_var.get()).name}))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self._set_buttons("normal"))

    def _fill_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for image in self.images:
            size_kb = round(int(image.get("size_bytes", 0)) / 1024, 1)
            size_label = f"{size_kb} KB"
            iid = str(image.get("xref"))
            self.tree.insert("", "end", iid=iid, values=(image.get("page"), image.get("image_number"), size_label))

    def select_all(self) -> None:
        self.tree.selection_set(self.tree.get_children())

    def clear_selection(self) -> None:
        self.tree.selection_remove(self.tree.selection())
        self.preview_label.configure(image=None, text="Select an image")
        self.preview_image = None

    def show_selected_preview(self, event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        xref = int(selection[0])
        try:
            image_bytes, _ext = get_image_bytes(Path(self.pdf_var.get()), xref)
            image = Image.open(BytesIO(image_bytes))
            # Large preview: preserve aspect ratio but use more of the panel.
            image.thumbnail((620, 460))
            self.preview_image = ctk.CTkImage(light_image=image, dark_image=image, size=image.size)
            self.preview_label.configure(image=self.preview_image, text="")
        except Exception as exc:
            self.preview_label.configure(image=None, text=f"Preview failed:\n{exc}")
            self.preview_image = None

    def start_extract(self) -> None:
        if not self._validate_pdf():
            return
        selected = [int(iid) for iid in self.tree.selection()]
        if not selected:
            messagebox.showerror("No images selected", "Please select one or more images from the table.")
            return
        self.clear_log()
        self._set_buttons("disabled")
        threading.Thread(target=self._extract_worker, args=(selected,), daemon=True).start()

    def _extract_worker(self, selected: list[int]) -> None:
        try:
            pdf_path = Path(self.pdf_var.get())
            output_dir = Path(self.output_var.get())
            out = extract_selected_images_from_pdf(
                pdf_path=pdf_path,
                output_dir=output_dir,
                xrefs=selected,
                log=self.thread_safe_log,
                progress=self._update_progress,
            )
            report = output_dir / "paperkit_image_extraction_report.xlsx"
            export_operation_report(
                report,
                "PaperKit Image Extraction Report",
                {
                    "PDF": pdf_path.name,
                    "Images found": len(self.images),
                    "Images selected": len(selected),
                    "Output folder": str(out),
                },
                [img for img in self.images if int(img.get("xref")) in set(selected)],
                log=self.thread_safe_log,
            )
            self.after(0, lambda: self._show_summary("Images extracted", {"Selected": len(selected), "Folder": out.name, "Report": report.name}))
            self.after(0, lambda: messagebox.showinfo("Done", f"Selected images extracted:\n{out}\n\nReport created:\n{report}"))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self._set_buttons("normal"))
