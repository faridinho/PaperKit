# Ghostscript Setup

PaperKit uses Ghostscript only for PDF compression.

Rename, metadata export, citation export, duplicate detection, and image extraction do not require Ghostscript.

## Download

Download Ghostscript from:

https://ghostscript.com/releases/gsdnld.html

For most Windows users, install the 64-bit Windows version.

## After Installation

Restart PaperKit after installing Ghostscript.

If compression still does not work, paste the full path to `gswin64c.exe` into the Ghostscript executable field.

Common Windows path example:

```text
C:\Program Files\gs\gs10.xx.x\bin\gswin64c.exe
```

## Compression Profiles

PaperKit supports these Ghostscript quality profiles:

- `screen`: smallest file size, lower quality
- `ebook`: balanced quality and size
- `printer`: higher quality
- `prepress`: highest quality, larger output
