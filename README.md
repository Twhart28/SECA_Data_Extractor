# SECA Data Converter

A helper script that extracts patient metadata and measurement values from SECA PDF reports and exports them into an Excel workbook.

## Setup
1. Create and activate a virtual environment.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   The requirements list includes `pyarrow`, which prevents pandas from failing during import on Windows.
3. Ensure Tesseract OCR is installed. The script expects it at `C:\\Program Files\\Tesseract-OCR\\tesseract.exe`; update `pytesseract.pytesseract.tesseract_cmd` in `seca_data_converter.py` if it lives elsewhere.

## Running
From an activated environment, launch the converter GUI with:
```bash
python seca_data_converter.py
```
Select the PDF files and destination folder when prompted, and the script will write an Excel file containing the extracted data.