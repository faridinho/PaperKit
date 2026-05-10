# Building PaperKit for Windows

A folder-based PyInstaller build is recommended because it is more reliable than a single-file executable.

## 1. Create a clean virtual environment

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## 2. Build the app

Run:

```bat
build_windows_folder.bat
```

The build will be created in:

```text
dist\PaperKit\
```

## 3. Test the app

Run:

```text
dist\PaperKit\PaperKit.exe
```

## 4. Distribute the app

Zip the whole folder:

```text
dist\PaperKit\
```

Do not distribute only `PaperKit.exe` from the folder build. The app needs the nearby bundled files and folders.

## Notes

- Compression requires Ghostscript on the user's computer.
- Other features work without Ghostscript.
- Unsigned Windows apps may show a SmartScreen warning.
