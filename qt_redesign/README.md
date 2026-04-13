# SECA Data Converter Qt Redesign

This folder contains a standalone PySide6/Qt version of the SECA Data Converter. The Qt app now has its own local backend in `qt_redesign/backend.py` and does not rely on `seca_data_converter.py`.

## Run locally

From the repository root:

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\qt_redesign\requirements.txt
.\.venv\Scripts\python.exe .\qt_redesign\app.py
```

Tesseract OCR still needs to be installed locally. The current backend expects the Windows default path:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

## Workflow

1. Add one or more SECA PDF reports.
2. Choose the Excel export path.
3. Process the reports.
4. Review flagged OCR fields first.
5. Optionally edit all extracted rows.
6. Export the same `All Data` workbook structure as the original app.

## Standalone files

- `app.py`: Qt user interface
- `backend.py`: PDF parsing, OCR, QC rules, and Excel export logic
- `App_Logo.ico`: local icon used by the Qt app and PyInstaller spec
- `seca_qt_converter.spec`: PyInstaller build spec for the standalone Qt executable

## Build an executable

Install PyInstaller in the environment if needed:

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

Then run from the repository root:

```powershell
.\.venv\Scripts\pyinstaller.exe .\qt_redesign\seca_qt_converter.spec
```

The executable will be created under `dist\seca_qt_converter`.
