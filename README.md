# PaperKit

Academic PDF workflow toolkit for renaming, compressing, metadata export, citation export, duplicate detection, and image extraction.

## Run from source

```bash
pip install -r requirements.txt
python app.py
```

## Build Windows folder app

```bat
build_windows_folder.bat
```

Share the whole `dist/PaperKit` folder as a ZIP.

## Notes

Compression requires Ghostscript. Rename, metadata export, citation export, duplicate detection, and image extraction do not require Ghostscript.

## License

PaperKit is recommended to be released under the MIT License. It is permissive and suitable for a free academic utility intended to help researchers and fellow scientists.

## Reports

PaperKit exports Excel operation reports for major workflows so users can verify the results after processing PDFs.


## Drag-and-drop

PaperKit supports drag-and-drop when `tkinterdnd2` is installed. You can drag PDF folders into folder fields and drag a single PDF into the Image Extraction PDF field.
