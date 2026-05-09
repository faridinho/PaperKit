from __future__ import annotations

import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from pdf_tools import (
    PDF_SETTINGS,
    extract_images_from_pdf,
    run_compress_only,
    run_extract_metadata,
    run_rename_and_compress,
    run_rename_only,
)

APP_TITLE = "Academic PDF Tools"
PROGRAMMER_NAME = "Farid Gazani"
GITHUB_PAGE = "https://github.com/faridinho/PaperKit"


class PDFToolsApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("980x800")
        self.minsize(880, 680)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Batch-tool variables
        self.input_var = ctk.StringVar()
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.mode_var = ctk.StringVar(value="both")
        self.quality_var = ctk.StringVar(value="ebook")
        self.start_var = ctk.StringVar(value="1")
        self.gs_var = ctk.StringVar(value="")
        self.in_place_var = ctk.BooleanVar(value=False)
        self.keep_renamed_var = ctk.BooleanVar(value=False)
        self.lowercase_rest_var = ctk.BooleanVar(value=False)

        # Image-extractor variables
        self.image_pdf_var = ctk.StringVar()
        self.image_output_var = ctk.StringVar(value=str(Path.cwd() / "extracted_images"))

        self._build_ui()
        self._update_option_states()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        title = ctk.CTkLabel(self, text=APP_TITLE, font=ctk.CTkFont(size=26, weight="bold"))
        title.grid(row=0, column=0, padx=20, pady=(18, 8), sticky="w")

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=1, column=0, padx=20, pady=(0, 12), sticky="nsew")

        self.batch_tab = self.tabs.add("Batch PDF tools")
        self.image_tab = self.tabs.add("Image extractor")

        self._build_batch_tab(self.batch_tab)
        self._build_image_tab(self.image_tab)

        footer_frame = ctk.CTkFrame(self)
        footer_frame.grid(row=2, column=0, padx=20, pady=(0, 14), sticky="ew")
        footer_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(footer_frame, text=f"Programmer: {PROGRAMMER_NAME}", anchor="w").grid(
            row=0, column=0, padx=12, pady=10, sticky="w"
        )
        ctk.CTkButton(footer_frame, text="Open GitHub page", command=self.open_github_page).grid(
            row=0, column=1, padx=12, pady=10, sticky="e"
        )

    def _build_batch_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(5, weight=1)

        folder_frame = ctk.CTkFrame(parent)
        folder_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        folder_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folder_frame, text="Input folder").grid(row=0, column=0, padx=12, pady=12, sticky="w")
        ctk.CTkEntry(folder_frame, textvariable=self.input_var).grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        ctk.CTkButton(folder_frame, text="Browse", command=self.choose_input).grid(row=0, column=2, padx=12, pady=12)

        ctk.CTkLabel(folder_frame, text="Output folder").grid(row=1, column=0, padx=12, pady=12, sticky="w")
        ctk.CTkEntry(folder_frame, textvariable=self.output_var).grid(row=1, column=1, padx=8, pady=12, sticky="ew")
        ctk.CTkButton(folder_frame, text="Browse", command=self.choose_output).grid(row=1, column=2, padx=12, pady=12)

        mode_frame = ctk.CTkFrame(parent)
        mode_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        mode_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(mode_frame, text="Mode", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")
        for col, (label, value) in enumerate([
            ("Rename only", "rename"),
            ("Compress only", "compress"),
            ("Rename + Compress", "both"),
            ("Extract metadata", "metadata"),
        ]):
            ctk.CTkRadioButton(
                mode_frame,
                text=label,
                variable=self.mode_var,
                value=value,
                command=self._update_option_states,
            ).grid(row=1, column=col, padx=12, pady=(6, 14), sticky="w")

        options_frame = ctk.CTkFrame(parent)
        options_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        options_frame.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(options_frame, text="Start number").grid(row=0, column=0, padx=12, pady=12, sticky="w")
        self.start_entry = ctk.CTkEntry(options_frame, textvariable=self.start_var, width=120)
        self.start_entry.grid(row=0, column=1, padx=8, pady=12, sticky="w")

        ctk.CTkLabel(options_frame, text="Compression quality").grid(row=0, column=2, padx=12, pady=12, sticky="w")
        self.quality_menu = ctk.CTkOptionMenu(options_frame, variable=self.quality_var, values=list(PDF_SETTINGS.keys()))
        self.quality_menu.grid(row=0, column=3, padx=8, pady=12, sticky="w")

        ctk.CTkLabel(options_frame, text="Ghostscript executable").grid(row=1, column=0, padx=12, pady=12, sticky="w")
        self.gs_entry = ctk.CTkEntry(options_frame, textvariable=self.gs_var, placeholder_text="Optional: gswin64c, gs, or full path")
        self.gs_entry.grid(row=1, column=1, columnspan=3, padx=8, pady=12, sticky="ew")

        help_text = (
            "Ghostscript can be left empty if gswin64c/gs is already in PATH. "
            "Metadata export creates an Excel file with ID, Title, Author, and Year. "
            "The ID is read from renamed files such as 1-Title.pdf."
        )
        self.help_label = ctk.CTkLabel(options_frame, text=help_text, wraplength=820, justify="left")
        self.help_label.grid(row=2, column=0, columnspan=4, padx=12, pady=(0, 12), sticky="w")

        checks_frame = ctk.CTkFrame(parent)
        checks_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        checks_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.in_place_check = ctk.CTkCheckBox(checks_frame, text="Rename originals directly", variable=self.in_place_var)
        self.in_place_check.grid(row=0, column=0, padx=12, pady=12, sticky="w")

        self.keep_check = ctk.CTkCheckBox(checks_frame, text="Keep renamed copies", variable=self.keep_renamed_var)
        self.keep_check.grid(row=0, column=1, padx=12, pady=12, sticky="w")

        self.lowercase_check = ctk.CTkCheckBox(checks_frame, text="Old title style: lowercase rest", variable=self.lowercase_rest_var)
        self.lowercase_check.grid(row=0, column=2, padx=12, pady=12, sticky="w")

        action_frame = ctk.CTkFrame(parent)
        action_frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)

        self.start_button = ctk.CTkButton(action_frame, text="Start", command=self.start_processing, height=40)
        self.start_button.grid(row=0, column=1, padx=12, pady=12)

        self.clear_button = ctk.CTkButton(action_frame, text="Clear log", command=self.clear_log, height=40)
        self.clear_button.grid(row=0, column=2, padx=12, pady=12)

        self.log_box = ctk.CTkTextbox(parent, wrap="word")
        self.log_box.grid(row=5, column=0, padx=10, pady=(10, 10), sticky="nsew")
        self.log("Choose folders, select a mode, then click Start.")

    def _build_image_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        file_frame = ctk.CTkFrame(parent)
        file_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="Single PDF file").grid(row=0, column=0, padx=12, pady=12, sticky="w")
        ctk.CTkEntry(file_frame, textvariable=self.image_pdf_var).grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        ctk.CTkButton(file_frame, text="Browse", command=self.choose_image_pdf).grid(row=0, column=2, padx=12, pady=12)

        ctk.CTkLabel(file_frame, text="Output folder").grid(row=1, column=0, padx=12, pady=12, sticky="w")
        ctk.CTkEntry(file_frame, textvariable=self.image_output_var).grid(row=1, column=1, padx=8, pady=12, sticky="ew")
        ctk.CTkButton(file_frame, text="Browse", command=self.choose_image_output).grid(row=1, column=2, padx=12, pady=12)

        info_frame = ctk.CTkFrame(parent)
        info_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        info_text = (
            "This extracts embedded raster images from one PDF using the original image bytes stored inside the file. "
            "It does not resize, downsample, or recompress the images, so this is the highest quality available from the PDF. "
            "Some figures may be vector drawings; those may not appear as extractable images."
        )
        ctk.CTkLabel(info_frame, text=info_text, wraplength=840, justify="left").grid(row=0, column=0, padx=12, pady=12, sticky="w")

        action_frame = ctk.CTkFrame(parent)
        action_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)

        self.image_start_button = ctk.CTkButton(action_frame, text="Extract images", command=self.start_image_extraction, height=40)
        self.image_start_button.grid(row=0, column=1, padx=12, pady=12)

        self.image_clear_button = ctk.CTkButton(action_frame, text="Clear log", command=self.clear_image_log, height=40)
        self.image_clear_button.grid(row=0, column=2, padx=12, pady=12)

        self.image_log_box = ctk.CTkTextbox(parent, wrap="word")
        self.image_log_box.grid(row=3, column=0, padx=10, pady=(10, 10), sticky="nsew")
        self.image_log("Choose one PDF, choose an output folder, then click Extract images.")

    def choose_input(self) -> None:
        folder = filedialog.askdirectory(title="Choose input folder")
        if folder:
            self.input_var.set(folder)

    def choose_output(self) -> None:
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_var.set(folder)

    def choose_image_pdf(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Choose one PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if file_path:
            self.image_pdf_var.set(file_path)

    def choose_image_output(self) -> None:
        folder = filedialog.askdirectory(title="Choose image output folder")
        if folder:
            self.image_output_var.set(folder)

    def open_github_page(self) -> None:
        webbrowser.open(GITHUB_PAGE)

    def _update_option_states(self) -> None:
        mode = self.mode_var.get()
        rename_related = mode in {"rename", "both"}
        compress_related = mode in {"compress", "both"}

        self.in_place_check.configure(state="normal" if mode == "rename" else "disabled")
        self.keep_check.configure(state="normal" if mode == "both" else "disabled")
        self.lowercase_check.configure(state="normal" if rename_related else "disabled")
        self.start_entry.configure(state="normal" if rename_related else "disabled")
        self.quality_menu.configure(state="normal" if compress_related else "disabled")
        self.gs_entry.configure(state="normal" if compress_related else "disabled")

    def log(self, message: str) -> None:
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def thread_safe_log(self, message: str) -> None:
        self.after(0, lambda: self.log(message))

    def clear_log(self) -> None:
        self.log_box.delete("1.0", "end")

    def image_log(self, message: str) -> None:
        self.image_log_box.insert("end", message + "\n")
        self.image_log_box.see("end")
        self.update_idletasks()

    def thread_safe_image_log(self, message: str) -> None:
        self.after(0, lambda: self.image_log(message))

    def clear_image_log(self) -> None:
        self.image_log_box.delete("1.0", "end")

    def start_processing(self) -> None:
        if not self.input_var.get().strip():
            messagebox.showerror("Missing input", "Please choose an input folder.")
            return
        if not self.output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return

        start_number = 1
        if self.mode_var.get() in {"rename", "both"}:
            try:
                start_number = int(self.start_var.get())
            except ValueError:
                messagebox.showerror("Invalid start number", "Start number must be an integer.")
                return

        self.start_button.configure(state="disabled", text="Working...")
        self.clear_log()

        worker = threading.Thread(target=self._run_worker, args=(start_number,), daemon=True)
        worker.start()

    def _run_worker(self, start_number: int) -> None:
        try:
            input_dir = Path(self.input_var.get()).expanduser().resolve()
            output_dir = Path(self.output_var.get()).expanduser().resolve()
            mode = self.mode_var.get()
            gs = self.gs_var.get().strip() or None

            if mode == "rename":
                run_rename_only(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    start_number=start_number,
                    in_place=self.in_place_var.get(),
                    lowercase_rest=self.lowercase_rest_var.get(),
                    log=self.thread_safe_log,
                )
            elif mode == "compress":
                run_compress_only(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    quality=self.quality_var.get(),
                    gs_executable=gs,
                    log=self.thread_safe_log,
                )
            elif mode == "metadata":
                run_extract_metadata(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    log=self.thread_safe_log,
                )
            else:
                run_rename_and_compress(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    start_number=start_number,
                    quality=self.quality_var.get(),
                    gs_executable=gs,
                    keep_renamed_copies=self.keep_renamed_var.get(),
                    lowercase_rest=self.lowercase_rest_var.get(),
                    log=self.thread_safe_log,
                )

            self.after(0, lambda: messagebox.showinfo("Done", "Processing finished."))
        except Exception as exc:
            self.thread_safe_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self.start_button.configure(state="normal", text="Start"))

    def start_image_extraction(self) -> None:
        if not self.image_pdf_var.get().strip():
            messagebox.showerror("Missing PDF", "Please choose a single PDF file.")
            return
        if not self.image_output_var.get().strip():
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return

        self.image_start_button.configure(state="disabled", text="Extracting...")
        self.clear_image_log()

        worker = threading.Thread(target=self._run_image_worker, daemon=True)
        worker.start()

    def _run_image_worker(self) -> None:
        try:
            pdf_path = Path(self.image_pdf_var.get()).expanduser().resolve()
            output_dir = Path(self.image_output_var.get()).expanduser().resolve()
            extract_images_from_pdf(
                pdf_path=pdf_path,
                output_dir=output_dir,
                log=self.thread_safe_image_log,
            )
            self.after(0, lambda: messagebox.showinfo("Done", "Image extraction finished."))
        except Exception as exc:
            self.thread_safe_image_log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self.image_start_button.configure(state="normal", text="Extract images"))


if __name__ == "__main__":
    app = PDFToolsApp()
    app.mainloop()
