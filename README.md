# SECA Data Converter

SECA Data Converter is a desktop utility for extracting patient metadata and body-composition measurements from SECA PDF reports and exporting them to Excel.

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
.\.venv\Scripts\pyinstaller.exe .\qt_redesign\seca_qt_converter.spec
```

The packaged app is created under `dist\seca_qt_converter`.

## Tesseract requirement

Tesseract OCR must be installed locally. The current backend expects the Windows default path:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

## Legacy Version 1

The previous Tkinter-based release is preserved in [`legacy_v1/`](./legacy_v1/) for reference and rollback. It is no longer the primary app in this repository.

## Repository

- GitHub: [Twhart28/SECA_Data_Converter](https://github.com/Twhart28/SECA_Data_Converter)
- Contact: `thomaswhart28@gmail.com`
