# SECA Data Converter

**Contact:** thomaswhart28@gmail.com

A Tkinter-based utility that extracts patient metadata and measurement values from SECA PDF reports and exports them to an Excel workbook. The script performs OCR on the report snapshots, validates common calculations, and lets you manually correct fields before saving.

## Prerequisites
- Python 3.9+ (Tkinter is bundled with most installations)
- Tesseract OCR installed locally
  - Windows default path: `C:\\Program Files\\Tesseract-OCR\\tesseract.exe`
  - If Tesseract lives elsewhere, update `pytesseract.pytesseract.tesseract_cmd` near the top of `seca_data_converter.py`.
- System packages required by `pytesseract`/Pillow to process images (platform-specific)

## Installation
1. Create and activate a virtual environment.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. (Optional) Verify Tesseract is discoverable:
   ```bash
   tesseract --version
   ```

## Running the converter
Launch the GUI from an activated environment:
```bash
python seca_data_converter.py
```
You will see a startup screen with two actions:
- **Continue**: proceed to file selection and conversion.
- **Open README**: view this guide inside the app.

### Typical workflow
1. **Select reports**: Choose one or more SECA PDF files when prompted.
2. **Choose output**: Pick a destination and filename for the Excel export (a timestamped suggestion is provided).
3. **Processing**: A progress window shows which file is being parsed.
4. **Review OCR (if needed)**: If blank fields or quality-control checks fail, the editor lists the affected fields with an OCR snapshot so you can correct the values.
5. **Export**: The script writes `All Data` to the Excel workbook, centers text cells, and confirms success.

### Output details
- Each row corresponds to one PDF.
- Columns include patient metadata, measurement values, calculated checks, and any QC flags.
- Rows that could not be recognized as SECA exports are placed at the top with a "Not recognized" message.

## Troubleshooting
- **Missing Tesseract**: Install it and update `pytesseract.pytesseract.tesseract_cmd` if necessary.
- **Excel export errors**: Ensure `openpyxl` is installed (included in `requirements.txt`).
- **Blank or incorrect OCR values**: Use the review dialog to correct them before export.

## Project files
- `seca_data_converter.py`: Main application logic and GUI.
- `requirements.txt`: Python dependencies.
- `App_Logo.ico`: Optional window icon.