from __future__ import annotations

import webbrowser

import customtkinter as ctk

from pdf_tools import GHOSTSCRIPT_DOWNLOAD_URL
from tabs.common import SURFACE, SURFACE_ALT, BORDER, TEXT, MUTED, PRIMARY, PRIMARY_HOVER


class HelpTab(ctk.CTkFrame):
    def __init__(
        self,
        master,
        app_title: str,
        app_version: str,
        github_page: str,
    ) -> None:
        super().__init__(master, fg_color="transparent")
        self.app_title = app_title
        self.app_version = app_version
        self.github_page = github_page
        self._build()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        scroll.grid_columnconfigure(0, weight=1)

        intro = ctk.CTkFrame(
            scroll,
            corner_radius=18,
            fg_color=SURFACE,
            border_width=1,
            border_color=BORDER,
        )
        intro.grid(row=0, column=0, padx=8, pady=(8, 14), sticky="ew")
        intro.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            intro,
            text=f"{self.app_title} Help",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=TEXT,
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(18, 4), sticky="ew")

        ctk.CTkLabel(
            intro,
            text=(
                "Welcome to PaperKit, an academic PDF workflow tool designed to help you manage "
                "research papers more easily. Use this page to learn what each tool does, when to use it, "
                "and what to check before exporting or changing files."
            ),
            text_color=MUTED,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=1, column=0, padx=18, pady=(0, 10), sticky="ew")

        ctk.CTkLabel(
            intro,
            text=(
                "PaperKit is designed to be safe by default. It does not change your original files "
                "unless you specifically enable an option that directly modifies them."
            ),
            text_color=MUTED,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=2, column=0, padx=18, pady=(0, 18), sticky="ew")

        sections = [
            (
                "Rename PDFs",
                "Use Rename to create clearer PDF filenames from extracted paper titles. Choose an input folder, "
                "choose an output folder, then click Preview renaming. Review the preview table carefully before "
                "clicking Apply renaming. If Add numbering is enabled, filenames will look like 1-Paper title.pdf, "
                "2-Another paper title.pdf, and so on. You can choose the starting number. If Rename originals directly "
                "is disabled, PaperKit creates renamed copies and keeps your original files unchanged. If it is enabled, "
                "PaperKit renames files in the original folder, so use it only after checking the preview."
            ),
            (
                "Rename preview table",
                "The preview table shows the old filename, the extracted title, the proposed new filename, and the status. "
                "This step is important because PDF title extraction is not always perfect. If a PDF has poor metadata or "
                "unusual first-page formatting, the extracted title may need checking before you apply the rename."
            ),
            (
                "Duplicate detection",
                "Duplicate detection is available in the Rename section. Use Detect duplicates to check whether a folder "
                "contains identical PDF files. PaperKit compares file content, not only filenames, so it can detect files "
                "that have different names but are actually the same PDF. The duplicate report is saved as an Excel file "
                "in your output folder. PaperKit does not delete duplicate files automatically. Review the report before "
                "removing anything."
            ),
            (
                "Compress PDFs",
                "Use Compress to reduce PDF file size. Choose the input folder, choose the output folder, select a compression "
                "quality, then start compression. The screen option makes the smallest files with lower quality. ebook is usually "
                "the best balance for sharing and archiving. printer and prepress keep higher quality but create larger files. "
                "The Cancel button stops the job as safely as possible and prevents the next files from starting."
            ),
            (
                "Ghostscript requirement",
                "PDF compression requires Ghostscript. Rename, metadata export, citation export, duplicate detection, and image "
                "extraction do not require Ghostscript. If compression does not work, install Ghostscript, restart PaperKit, "
                "and try again. If PaperKit still cannot find it, paste the full path to gswin64c.exe into the Ghostscript field."
            ),
            (
                "Metadata export",
                "Use Metadata to export PDF information into an Excel file. You can choose which fields to include, such as ID, "
                "Filename, Title, Author, Year, DOI, Page count, and File size. The ID is usually taken from filenames such as "
                "1-Paper title.pdf. Metadata quality depends on the PDF file. Some PDFs contain clean metadata, while others may "
                "have missing or incorrect title, author, year, or DOI information. Always review the exported Excel file before "
                "using it for formal academic records."
            ),
            (
                "Citation export",
                "Use Citations to create RIS and BibTeX files from your PDF collection. RIS files can be imported into tools such "
                "as Mendeley, EndNote, Zotero, and RefWorks. BibTeX files are useful for LaTeX, JabRef, Zotero, and Mendeley. "
                "Citation files are generated automatically from metadata PaperKit can extract from the PDFs. Citation quality "
                "depends on the quality of the PDF metadata, so imported records should always be reviewed in your reference manager."
            ),
            (
                "Image Extraction",
                "Use Image Extraction to extract images from a single PDF. Choose one PDF file, click Scan images, review the image "
                "list and preview, select the images you want, then click Extract selected images. PaperKit extracts embedded images "
                "from the PDF at their original stored quality. It does not resize, downsample, or recompress them. Some academic "
                "figures are vector graphics or page drawings rather than embedded raster images. These may be visible in the PDF "
                "but may not appear in the image list."
            ),
            (
                "Output folders and reports",
                "Most tools save results in the output folder you choose. Depending on the tool, PaperKit may create renamed PDFs, "
                "compressed PDFs, metadata Excel files, duplicate reports, RIS citation files, BibTeX citation files, extracted images, "
                "and operation reports. Use Open output folder after a task finishes to quickly view the results."
            ),
            (
                "Common issue: extracted title is wrong",
                "PDF title extraction is not perfect. Always use Preview renaming before applying changes. If a PDF has poor metadata, "
                "a scanned first page, or unusual formatting, PaperKit may choose the wrong title. The preview gives you a chance to "
                "catch problems before filenames are changed."
            ),
            (
                "Common issue: metadata or citations are incomplete",
                "Many PDFs do not contain complete metadata. PaperKit extracts the best information it can find, but some records may "
                "be missing authors, years, journals, or DOI values. Review Excel metadata and imported citation records before using "
                "them in a thesis, paper, or reference library."
            ),
            (
                "Common issue: images are missing from the image list",
                "Some figures are not stored as normal images inside a PDF. They may be vector graphics, charts, or drawings made from "
                "PDF instructions. PaperKit extracts embedded raster images only, so vector graphics may not appear in the image list."
            ),
            (
                "About PaperKit",
                f"PaperKit version {self.app_version} is built to support researchers, students, and scientists who work with large "
                "collections of academic PDFs. It is intended to make common PDF tasks faster, safer, and easier."
            ),
        ]

        row = 1
        for title, body in sections:
            self._add_help_card(scroll, row, title, body)
            row += 1

        self._add_ghostscript_card(scroll, row)

    def _add_help_card(self, parent: ctk.CTkFrame, row: int, title: str, body: str) -> None:
        card = ctk.CTkFrame(
            parent,
            corner_radius=16,
            fg_color=SURFACE_ALT,
            border_width=1,
            border_color=BORDER,
        )
        card.grid(row=row, column=0, padx=8, pady=8, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=TEXT,
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(16, 6), sticky="ew")

        ctk.CTkLabel(
            card,
            text=body,
            text_color=MUTED,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=1, column=0, padx=18, pady=(0, 16), sticky="ew")

    def _add_ghostscript_card(self, parent: ctk.CTkFrame, row: int) -> None:
        card = ctk.CTkFrame(
            parent,
            corner_radius=16,
            fg_color=SURFACE_ALT,
            border_width=1,
            border_color=BORDER,
        )
        card.grid(row=row, column=0, padx=8, pady=(8, 18), sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card,
            text="Download Ghostscript for compression",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=TEXT,
            anchor="w",
        ).grid(row=0, column=0, padx=18, pady=(16, 6), sticky="ew")

        ctk.CTkLabel(
            card,
            text=(
                "Only the Compress tool needs Ghostscript. Install the 64-bit Windows version, then restart PaperKit. "
                "If compression still does not work, paste the full path to gswin64c.exe into the Ghostscript field."
            ),
            text_color=MUTED,
            wraplength=900,
            justify="left",
            anchor="w",
        ).grid(row=1, column=0, padx=18, pady=(0, 10), sticky="ew")

        ctk.CTkButton(
            card,
            text="Open Ghostscript download page",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            command=lambda: webbrowser.open(GHOSTSCRIPT_DOWNLOAD_URL),
        ).grid(row=2, column=0, padx=18, pady=(0, 18), sticky="w")
