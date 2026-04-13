# SECA Data Extractor Version 2

This folder contains the standalone **Version 2** PySide6/Qt application. The Qt app has its own local backend in `qt_redesign/backend.py` and does not rely on the older Tkinter implementation in `legacy_v1/`.

## Run locally

From the repository root:

```powershell
.\.venv\Scripts\python.exe -m pip install -r .\qt_redesign\requirements.txt
.\.venv\Scripts\python.exe .\qt_redesign\app.py
```

For source-checkout runs, the backend can use a local Tesseract install. The default Windows path is:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

You can also override that with:

```powershell
$env:SECA_TESSERACT_CMD="C:\Path\To\tesseract.exe"
```

or:

```powershell
$env:SECA_TESSERACT_DIR="C:\Path\To\Tesseract-OCR"
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
- `seca_data_extractor.spec`: PyInstaller build spec for the standalone Qt executable
- `seca_data_extractor.iss`: Inno Setup installer definition
- `package_release.ps1`: release helper that builds the portable zip and installer

## Build an executable

Install PyInstaller in the environment if needed:

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

Then run from the repository root:

```powershell
.\.venv\Scripts\pyinstaller.exe .\qt_redesign\seca_data_extractor.spec
```

The executable will be created under `dist\seca_data_extractor`.

The PyInstaller spec bundles the OCR runtime from `SECA_TESSERACT_DIR` if set, otherwise from:

```text
C:\Program Files\Tesseract-OCR
```

## Build a handoff package

From the repository root:

```powershell
.\qt_redesign\package_release.ps1
```

That script creates:
- `dist\seca_data_extractor\` portable app folder
- `dist\seca_data_extractor_portable.zip` for direct sharing
- `dist\seca_data_extractor_setup.exe` if Inno Setup 6 is installed on the build machine
