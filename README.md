# MarkItDown Handy

**Version:** 0.1.0

MarkItDown Handy is a small macOS GUI for batch-converting files to Markdown with [Microsoft MarkItDown](https://github.com/microsoft/markitdown). It is designed for quick local conversion, scanned PDF OCR fallback, and simple release packaging.

## Features

- Add individual files or a folder of files.
- Drag-and-drop support when `tkinterdnd2` is available.
- Convert common inputs supported by MarkItDown extras, including PDF, DOCX, PPTX, XLSX, HTML, TXT, CSV/TSV, JSON, XML, images, audio, ZIP, and EPUB.
- Automatically tries direct conversion first.
- Falls back to OCR for scanned or weak PDF output when enabled.
- Shows queue status, current step, progress bars, and live logs.
- Does **not** overwrite existing output files. Existing names are resolved as `name_1.md`, `name_2.md`, etc.
- Builds either a small conda-wrapper app or a fully portable app with embedded dependencies.

## Repository layout

```text
src/markitdown_handy.py                      Main Tkinter GUI
scripts/run_from_source.sh                   Run locally from source
scripts/setup_dev_env.sh                     Install/update local conda dependencies
scripts/build_conda_wrapper_app.sh           Build small app using existing conda env
scripts/build_portable_embedded_app.sh       Build portable app with embedded runtime
scripts/build_both_apps.sh                   Build both app variants
scripts/diagnose_portable_app.sh             Diagnose portable app dependencies
```

## Quick start from source

```bash
conda create -n markitdown python=3.11 -y
conda activate markitdown
pip install 'markitdown[all]' tkinterdnd2
./scripts/run_from_source.sh
```

If you already have a `markitdown` conda environment, run:

```bash
./scripts/setup_dev_env.sh
./scripts/run_from_source.sh
```

## Build the conda-wrapper app

This app is small, but the target Mac must have the `markitdown` conda environment and OCR tools installed.

```bash
chmod +x scripts/*.sh
./scripts/build_conda_wrapper_app.sh
```

Output:

```text
dist-conda/MarkItDown Handy.app
dist-conda/MarkItDown_Handy_v0.1.0_conda_wrapper_macOS_<arch>.zip
```

## Build the portable app

This app embeds Python, MarkItDown, OCRmyPDF, Tesseract, Ghostscript, qpdf, and related dependencies. The target Mac does not need conda or Homebrew, but it must match the CPU architecture used for building.

```bash
chmod +x scripts/*.sh
rm -rf dist-portable build-portable
./scripts/build_portable_embedded_app.sh
```

Output:

```text
dist-portable/MarkItDown Handy Portable.app
dist-portable/MarkItDown_Handy_v0.1.0_portable_macOS_<arch>.zip
```

## Diagnose the portable app

```bash
./scripts/diagnose_portable_app.sh "dist-portable/MarkItDown Handy Portable.app"
```

Expected output includes:

```text
markitdown OK
ocrmypdf OK
tkinterdnd2 OK
```

## Notes for GitHub release

- The portable app is architecture-specific: build separately for `arm64` and `x86_64` if you need both.
- The portable app can be large because it embeds a full runtime.
- For public distribution, sign and notarize the `.app` with an Apple Developer ID.
- This project does not modify MarkItDown; it only provides a local GUI wrapper and packaging scripts.

## License

MIT License. See [`LICENSE`](LICENSE).
