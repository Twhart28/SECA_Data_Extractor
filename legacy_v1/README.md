# SECA Data Converter

Archived **Version 1** of the application. This is the previous Tkinter-based implementation and is kept for reference.

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

## Program workflow (end-to-end)
1. **Startup screen**: Choose whether to continue or open the README.
2. **Select reports**: Choose one or more SECA PDF files when prompted.
3. **Choose output**: Pick a destination and filename for the Excel export (a timestamped suggestion is provided).
4. **Parse each PDF**:
   - The app reads the PDF text layer to confirm it looks like a SECA export (expects “Patient Data” and “Single Measurement”).
   - Patient metadata (ID, sex, age, collection date/time) is extracted from the text layer.
   - For numeric measurements, the app crops known regions of the PDF image and runs OCR.
5. **Compute calculated fields**:
   - `Body Mass Index (kg/m^2)` is calculated from Weight and Height.
6. **Run data-quality (QC) checks**:
   - QC checks are run for every file based on numeric values and expected relationships.
   - Each failure produces a QC code in `Data Quality Fails` and sets `Data Quality` to `Fail`.
7. **Review OCR when needed**:
   - If there are blank OCR fields or QC failures, you’ll be prompted to review.
   - The review dialog shows the OCR snapshot for each flagged field and lets you correct values.
   - After edits, QC checks are re-run for the updated rows.
8. **Export to Excel**:
   - Output is written to the `All Data` sheet, centered for readability.
   - Rows that are not recognized as SECA exports are placed at the top with `Not recognized as a SECA data export` in `Data Quality Fails`.

## Output details
- Each row corresponds to one PDF.
- Columns include patient metadata, measurement values, calculated checks, and any QC flags.
- Rows that could not be recognized as SECA exports are placed at the top with a "Not recognized" message.

## Data-quality (QC) checks and tolerances
The QC system compares related measurements to expected formulas. If any required fields are missing, the check fails. Tolerances below are absolute (not percent). The `Data Quality Fails` column will contain one or more codes when a check fails. Use the map below to diagnose which values likely need review.

| QC Code | Check | Expected Relationship | Tolerance |
| --- | --- | --- | --- |
| 1 | Fat Mass (kg) + Fat-Free Mass (kg) vs Weight (kg) | `Fat Mass (kg) + Fat-Free Mass (kg) = Weight (kg)` | ±0.01 |
| 2 | Fat Mass (%) + Fat-Free Mass (%) | `Fat Mass (%) + Fat-Free Mass (%) = 100` | ±0.01 |
| 3 | Fat Mass Index + Fat-Free Mass Index vs SECA BMI | `Fat Mass Index (kg/m^2) + Fat-Free Mass Index (kg/m^2) = SECA BMI (kg/m^2)` | ±0.02 |
| 4 | Segmental lean mass vs Skeletal Muscle Mass | `Right Arm + Left Arm + Right Leg + Left Leg + Torso = Skeletal Muscle Mass (kg)` | ±0.03 |
| 5 | Weight/Height vs SECA BMI | `Weight (kg) / Height (m)^2 = SECA BMI (kg/m^2)` | ±0.3 |
| 6 | ECW/TBW using liters | `Extracellular Water (L) / Total Body Water (L) * 100 = ECW/TBW (%)` | ±0.03 |
| 7 | ECW/TBW using percentages | `Extracellular Water (%) / Total Body Water (%) * 100 = ECW/TBW (%)` | ±0.02 |
| 8 | Energy consumption | `Resting Energy Expenditure (kcal/day) * Physical Activity Level = Energy Consumption (kcal/day)` | ±0.02 |
| 9 | Phase Angle from resistance/reactance | `atan(Reactance / Resistance) * 180 / π = Phase Angle (deg)` | ±0.1 |
| 10 | Phase Angle Percentile bounds | `0 <= Phase Angle Percentile <= 100` | Valid range |

### QC interpretation notes
- **Missing data fails the check**: if any required inputs are blank, the check fails and the code appears.
- **Phase Angle OCR correction**: for QC code **9**, if the calculated phase angle matches a version with a decimal inserted after the first digit (e.g., `41` → `4.1`), the value is auto-corrected and the check passes.
- **“Not recognized as a SECA data export”**: this is not a numerical QC check. It indicates the PDF doesn’t match the expected SECA text layout.

## Troubleshooting
- **Missing Tesseract**: Install it and update `pytesseract.pytesseract.tesseract_cmd` if necessary.
- **Excel export errors**: Ensure `openpyxl` is installed (included in `requirements.txt`).
- **Blank or incorrect OCR values**: Use the review dialog to correct them before export.

## Project files
- `seca_data_converter.py`: Main application logic and GUI.
- `requirements.txt`: Python dependencies.
- `App_Logo.ico`: Optional window icon.
