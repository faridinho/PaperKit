from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Callable

try:
    import fitz  # PyMuPDF: pip install pymupdf
except ImportError as exc:
    raise SystemExit("Missing dependency: install PyMuPDF with `pip install pymupdf`") from exc

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass

LogFn = Callable[[str], None]

SKIP_KEYWORDS = [
    "research article", "accepted", "submitted", "doi", "issn",
    "journal", "volume", "issue", "published online", "crossmark",
    "copyright", "special collection", "proceedings",
]

PDF_SETTINGS = {
    "screen": "/screen",      # lowest quality, smallest files
    "ebook": "/ebook",        # good default
    "printer": "/printer",    # higher quality
    "prepress": "/prepress",  # highest quality, larger files
}


def _log(message: str, log: LogFn | None = None) -> None:
    if log:
        log(message)
    else:
        print(message)


def clean_filename(text: str, max_length: int = 150, lowercase_rest: bool = False) -> str:
    """Return text that is safe to use as a filename."""
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip()

    if lowercase_rest:
        text = text.capitalize()

    return text[:max_length] or "Untitled"


def extract_title(pdf_path: Path, lowercase_rest: bool = False) -> str:
    """Extract a likely paper title from PDF metadata or the first page."""
    with fitz.open(pdf_path) as doc:
        metadata_title = (doc.metadata or {}).get("title")
        if metadata_title and len(metadata_title.strip()) > 10:
            return clean_filename(metadata_title, lowercase_rest=lowercase_rest)

        if doc.page_count == 0:
            return "Untitled"

        first_page_text = doc[0].get_text("text")

    lines = [line.strip() for line in first_page_text.split("\n") if len(line.strip()) > 10]
    filtered = [
        line for line in lines
        if not any(keyword in line.lower() for keyword in SKIP_KEYWORDS)
    ]

    candidate = max(filtered, key=lambda line: len(line.split()), default="Untitled")
    return clean_filename(candidate, lowercase_rest=lowercase_rest)


def unique_path(path: Path) -> Path:
    """Return a unique path by appending (2), (3), etc. if needed."""
    if not path.exists():
        return path

    base = path.with_suffix("")
    ext = path.suffix
    counter = 2

    while True:
        candidate = Path(f"{base} ({counter}){ext}")
        if not candidate.exists():
            return candidate
        counter += 1


def list_pdfs(folder: Path) -> list[Path]:
    """Return sorted PDF paths from a folder."""
    return sorted(path for path in folder.glob("*.pdf") if path.is_file())


def validate_folder(input_dir: Path) -> None:
    """Validate input folder and ensure it contains PDFs."""
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"Input folder does not exist: {input_dir}")

    if not list_pdfs(input_dir):
        raise ValueError(f"No PDF files found in: {input_dir}")


def prepare_input_files(input_dir: Path, work_dir: Path, copy_files: bool) -> list[Path]:
    """
    Return PDF paths for processing.

    If copy_files=True, PDFs are copied into work_dir first so originals are not modified.
    If copy_files=False, original PDFs are returned and may be renamed in place.
    """
    pdfs = list_pdfs(input_dir)

    if not copy_files:
        return pdfs

    work_dir.mkdir(parents=True, exist_ok=True)
    copied_paths: list[Path] = []

    for pdf_path in pdfs:
        destination = unique_path(work_dir / pdf_path.name)
        shutil.copy2(pdf_path, destination)
        copied_paths.append(destination)

    return copied_paths


def rename_pdfs(
    pdf_paths: list[Path],
    start_number: int = 1,
    lowercase_rest: bool = False,
    log: LogFn | None = None,
) -> list[Path]:
    """Rename PDFs using extracted titles and numbering."""
    renamed_paths: list[Path] = []
    counter = start_number

    for old_path in sorted(pdf_paths):
        title = extract_title(old_path, lowercase_rest=lowercase_rest)
        new_name = f"{counter}-{title}.pdf"
        new_path = unique_path(old_path.parent / new_name)

        if old_path.resolve() == new_path.resolve():
            _log(f"Skipped already named: {old_path.name}", log)
            renamed_paths.append(old_path)
        else:
            old_path.rename(new_path)
            _log(f"Renamed: {old_path.name} -> {new_path.name}", log)
            renamed_paths.append(new_path)

        counter += 1

    return renamed_paths


def find_ghostscript(user_gs: str | None = None) -> str:
    """Find a Ghostscript executable."""
    candidates = [user_gs] if user_gs else []
    candidates += ["gswin64c", "gswin32c", "gs"]

    for candidate in candidates:
        if candidate and shutil.which(candidate):
            return candidate

    raise RuntimeError(
        "Ghostscript was not found. Install Ghostscript and make sure gswin64c, "
        "gswin32c, or gs is on PATH. If it is not on PATH, enter the full path "
        "to gswin64c.exe in the Ghostscript executable field."
    )


def compress_pdf(
    pdf_path: Path,
    output_dir: Path,
    gs: str,
    quality: str = "ebook",
    log: LogFn | None = None,
) -> Path | None:
    """Compress one PDF into output_dir using Ghostscript."""
    if quality not in PDF_SETTINGS:
        raise ValueError(f"Unknown quality preset: {quality}")

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = unique_path(output_dir / pdf_path.name)

    command = [
        gs,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={PDF_SETTINGS[quality]}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={out_path}",
        str(pdf_path),
    ]

    result = subprocess.run(command, capture_output=True, text=True)

    if result.returncode != 0:
        _log(f"FAILED: {pdf_path.name}", log)
        if result.stderr:
            _log(result.stderr, log)
        return None

    original_mb = pdf_path.stat().st_size / 1024 / 1024
    compressed_mb = out_path.stat().st_size / 1024 / 1024
    reduction = 100 * (1 - compressed_mb / original_mb) if original_mb else 0

    _log(
        f"Compressed: {pdf_path.name}: "
        f"{original_mb:.2f} MB -> {compressed_mb:.2f} MB ({reduction:.1f}% smaller)",
        log,
    )
    return out_path


def _extract_id_from_filename(pdf_path: Path) -> str:
    """Extract the leading ID from names like 12-Title.pdf or 12_Title.pdf."""
    match = re.match(r"^\s*(\d+)\s*[-_]", pdf_path.stem)
    return match.group(1) if match else ""


def _extract_year_from_pdf(pdf_path: Path, metadata: dict, first_page_text: str) -> str:
    """
    Extract a likely publication year.

    PDF files do not always contain a true publication year. This function uses:
    1) metadata dates, then 2) first-page text, then 3) filename.
    """
    for key in ("creationDate", "modDate"):
        value = metadata.get(key) or ""
        match = re.search(r"(19|20)\d{2}", value)
        if match:
            return match.group(0)

    combined_text = f"{first_page_text}\n{pdf_path.stem}"
    years = re.findall(r"\b(?:19|20)\d{2}\b", combined_text)
    if years:
        return years[0]

    return ""


def extract_metadata_row(pdf_path: Path) -> dict[str, str]:
    """Extract ID, title, author, and likely publication year for one PDF."""
    first_page_text = ""
    metadata: dict = {}

    with fitz.open(pdf_path) as doc:
        metadata = doc.metadata or {}
        if doc.page_count:
            first_page_text = doc[0].get_text("text")

    title = metadata.get("title") or ""
    if not title or len(title.strip()) < 4:
        title = extract_title(pdf_path)

    author = (metadata.get("author") or "").strip()
    year = _extract_year_from_pdf(pdf_path, metadata, first_page_text)

    return {
        "ID": _extract_id_from_filename(pdf_path),
        "Title": re.sub(r"\s+", " ", title).strip(),
        "Author": author,
        "Year": year,
    }


def export_metadata_to_excel(rows: list[dict[str, str]], output_path: Path) -> Path:
    """Write metadata rows to an Excel workbook."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.worksheet.table import Table, TableStyleInfo
    except ImportError as exc:
        raise RuntimeError("Missing dependency: install openpyxl with `pip install openpyxl`") from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Metadata"

    headers = ["ID", "Title", "Author", "Year"]
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    widths = {"A": 10, "B": 60, "C": 35, "D": 12}
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A2"

    if rows:
        table_ref = f"A1:D{len(rows) + 1}"
        table = Table(displayName="PDFMetadataTable", ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    wb.save(output_path)
    return output_path


def run_extract_metadata(
    input_dir: Path,
    output_dir: Path,
    output_filename: str = "pdf_metadata.xlsx",
    log: LogFn | None = None,
) -> Path:
    """Extract selected PDF metadata into an Excel file."""
    validate_folder(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdfs = list_pdfs(input_dir)
    rows: list[dict[str, str]] = []

    _log("Mode: extract metadata", log)
    _log(f"Input folder: {input_dir}", log)

    for pdf_path in pdfs:
        row = extract_metadata_row(pdf_path)
        rows.append(row)
        display_id = row["ID"] or "no ID"
        _log(f"Metadata: {pdf_path.name} -> ID: {display_id}, Year: {row['Year'] or 'unknown'}", log)

    output_path = unique_path(output_dir / output_filename)
    export_metadata_to_excel(rows, output_path)
    _log(f"Metadata Excel saved in: {output_path}", log)
    return output_path


def run_rename_only(
    input_dir: Path,
    output_dir: Path,
    start_number: int = 1,
    in_place: bool = False,
    lowercase_rest: bool = False,
    log: LogFn | None = None,
) -> list[Path]:
    """Rename PDFs only."""
    validate_folder(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_paths = prepare_input_files(
        input_dir=input_dir,
        work_dir=output_dir / "renamed",
        copy_files=not in_place,
    )

    _log("Mode: rename only", log)
    _log(f"Input folder: {input_dir}", log)
    if in_place:
        _log("Warning: renaming original PDFs in place.", log)
    else:
        _log(f"Renamed PDFs will be saved in: {output_dir / 'renamed'}", log)

    return rename_pdfs(source_paths, start_number=start_number, lowercase_rest=lowercase_rest, log=log)


def run_compress_only(
    input_dir: Path,
    output_dir: Path,
    quality: str = "ebook",
    gs_executable: str | None = None,
    log: LogFn | None = None,
) -> list[Path]:
    """Compress PDFs only. Original PDFs are not modified."""
    validate_folder(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    gs = find_ghostscript(gs_executable)
    compressed_dir = output_dir / "compressed"
    compressed_paths: list[Path] = []

    _log("Mode: compress only", log)
    _log(f"Input folder: {input_dir}", log)
    _log(f"Compressed PDFs will be saved in: {compressed_dir}", log)
    _log(f"Ghostscript: {gs}\n", log)

    for pdf_path in list_pdfs(input_dir):
        compressed = compress_pdf(pdf_path, compressed_dir, gs=gs, quality=quality, log=log)
        if compressed:
            compressed_paths.append(compressed)

    return compressed_paths


def run_rename_and_compress(
    input_dir: Path,
    output_dir: Path,
    start_number: int = 1,
    quality: str = "ebook",
    gs_executable: str | None = None,
    keep_renamed_copies: bool = False,
    lowercase_rest: bool = False,
    log: LogFn | None = None,
) -> list[Path]:
    """Copy PDFs, rename the copies, compress them, and leave originals unchanged."""
    validate_folder(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    gs = find_ghostscript(gs_executable)
    renamed_dir = output_dir / "renamed"
    compressed_dir = output_dir / "compressed"
    compressed_paths: list[Path] = []

    _log("Mode: rename and compress", log)
    _log(f"Input folder: {input_dir}", log)
    _log(f"Output folder: {output_dir}", log)
    _log(f"Ghostscript: {gs}", log)
    _log("Original PDFs will not be modified.\n", log)

    copied_paths = prepare_input_files(input_dir, renamed_dir, copy_files=True)
    renamed_paths = rename_pdfs(copied_paths, start_number=start_number, lowercase_rest=lowercase_rest, log=log)

    _log("\nCompressing renamed PDFs...\n", log)
    for pdf_path in renamed_paths:
        compressed = compress_pdf(pdf_path, compressed_dir, gs=gs, quality=quality, log=log)
        if compressed:
            compressed_paths.append(compressed)

    if keep_renamed_copies:
        _log(f"\nRenamed uncompressed copies kept in: {renamed_dir}", log)
    else:
        shutil.rmtree(renamed_dir, ignore_errors=True)
        _log("\nTemporary renamed copies removed.", log)

    _log(f"Compressed PDFs saved in: {compressed_dir}", log)
    return compressed_paths


def extract_images_from_pdf(
    pdf_path: Path,
    output_dir: Path,
    log: LogFn | None = None,
) -> list[Path]:
    """
    Extract embedded images from a single PDF at the highest available quality.

    PyMuPDF returns the original embedded image bytes, so the images are not
    downsampled, resized, or recompressed by this function.
    """
    if not pdf_path.exists() or not pdf_path.is_file():
        raise ValueError(f"PDF file does not exist: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError(f"Selected file is not a PDF: {pdf_path}")

    target_dir = output_dir / f"{clean_filename(pdf_path.stem, max_length=80)}_images"
    target_dir.mkdir(parents=True, exist_ok=True)

    extracted_paths: list[Path] = []
    seen_xrefs: set[int] = set()

    _log("Mode: extract images", log)
    _log(f"PDF: {pdf_path}", log)
    _log(f"Output folder: {target_dir}", log)
    _log("Extracting original embedded image data without resizing or recompression.\n", log)

    with fitz.open(pdf_path) as doc:
        for page_index in range(doc.page_count):
            page = doc[page_index]
            images = page.get_images(full=True)

            if not images:
                _log(f"Page {page_index + 1}: no embedded images found", log)
                continue

            page_image_count = 0
            for image_index, image_info in enumerate(images, start=1):
                xref = image_info[0]

                # Avoid exporting the same embedded image multiple times when it is reused.
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)

                image_data = doc.extract_image(xref)
                image_bytes = image_data.get("image")
                extension = image_data.get("ext", "png")
                width = image_data.get("width", 0)
                height = image_data.get("height", 0)

                if not image_bytes:
                    continue

                filename = f"page_{page_index + 1:03d}_image_{image_index:03d}_xref_{xref}_{width}x{height}.{extension}"
                image_path = unique_path(target_dir / filename)
                image_path.write_bytes(image_bytes)
                extracted_paths.append(image_path)
                page_image_count += 1

                size_kb = image_path.stat().st_size / 1024
                _log(f"Extracted: {image_path.name} ({size_kb:.1f} KB)", log)

            if page_image_count == 0:
                _log(f"Page {page_index + 1}: only duplicate or unsupported images found", log)

    if extracted_paths:
        _log(f"\nDone. Extracted {len(extracted_paths)} unique images.", log)
        _log(f"Images saved in: {target_dir}", log)
    else:
        _log("\nNo extractable embedded images were found in this PDF.", log)
        _log("Note: some PDFs contain vector drawings or rendered page graphics instead of embedded raster images.", log)

    return extracted_paths
