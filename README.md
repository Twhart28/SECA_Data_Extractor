# SECA Data Extractor

SECA Data Extractor is a desktop utility for extracting patient metadata and body-composition measurements from SECA PDF reports and exporting them to Excel.

This repository now treats the **PySide6/Qt application as Version 2**, and keeps the older Tkinter implementation archived as **Version 1**.

## Current version

- **Version 2 (current):** [`qt_redesign/`](./qt_redesign/)
- **Version 1 (legacy):** [`legacy_v1/`](./legacy_v1/)

## Version 2 workflow

1. Add one or more SECA PDF reports.
2. Choose the Excel export path.
3. Process the reports locally.
4. Review flagged OCR fields first.
5. Optionally edit all extracted rows.
6. Export the Excel workbook.

## Version 2 features

- Standalone PySide6/Qt desktop app
- Local backend in `qt_redesign/backend.py`
- OCR snapshot review with guided correction flow
- In-app PDF viewer
- QC legend and structured overview table
- Standalone PyInstaller spec for packaging a separate executable

## Run Version 2 locally

From the repository root:

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\qt_redesign\requirements.txt
.\.venv\Scripts\python.exe .\qt_redesign\app.py
```

## Build Version 2 executable

```powershell
.\.venv\Scripts\pyinstaller.exe .\qt_redesign\seca_data_extractor.spec
```

The packaged app is created under `dist\seca_data_extractor`.

To create a handoff build for another user:

```powershell
.\qt_redesign\package_release.ps1
```

That script builds:
- a portable folder under `dist\seca_data_extractor`
- a portable zip under `dist\seca_data_extractor_portable.zip`
- an installer exe under `dist\seca_data_extractor_setup.exe` if Inno Setup is installed

## OCR runtime

The packaged Version 2 release bundles Tesseract OCR, so recipients do not need to install Python or Tesseract first.

For source-checkout runs, the app can use:
- a bundled runtime next to the exe
- `SECA_TESSERACT_CMD`
- `SECA_TESSERACT_DIR`
- a local Tesseract install on the Windows default path

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

## Legacy Version 1

The previous Tkinter-based release is preserved in [`legacy_v1/`](./legacy_v1/) for reference and rollback. It is no longer the primary app in this repository.

## Repository

- GitHub: [Twhart28/SECA_Data_Extractor](https://github.com/Twhart28/SECA_Data_Extractor)
- Contact: `thomaswhart28@gmail.com`
