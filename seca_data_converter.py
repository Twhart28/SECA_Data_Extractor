"""SECA Data Converter
=======================

This script opens GUI dialogs to let a user select one or more SECA PDF
reports and a destination folder.  It then extracts the patient metadata and
measurement values defined in the project requirements and stores them in an
Excel workbook (one row per PDF).

Usage::

    python seca_data_converter.py

Dependencies:
    - pdfplumber
    - pytesseract (requires the Tesseract OCR binary)
    - pandas (which also requires openpyxl for Excel output)
    - tkinter (bundled with most Python distributions)
"""

from __future__ import annotations

import math
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
import pytesseract
from PIL import ImageTk
from tkinter import Tk, messagebox
from tkinter import filedialog
from tkinter import ttk, StringVar

# --- OCR region configuration ---

# Base page size (in pixels) for the provided per-field coordinates
MEASUREMENT_BASE_WIDTH = 9917
MEASUREMENT_BASE_HEIGHT = 14034

# Render resolution for OCR snapshots (higher improves text clarity)
OCR_RENDER_RESOLUTION = 300

# Region definitions (left, top, right, bottom) in the base coordinate system.
# Each tuple pairs the destination field names with the corresponding crop box.
MEASUREMENT_CROP_BOXES = [
    (("Fat Mass (kg)", "Fat Mass (%)"), (7045, 3697, 8318, 4029)),
    (("Fat Mass Index (kg/m^2)",), (7045, 4029, 8318, 4543)),
    (("Fat-Free Mass (kg)", "Fat-Free Mass (%)"), (7045, 4543, 8318, 4857)),
    (("Fat-Free Mass Index (kg/m^2)",), (7045, 4857, 8318, 5365)),
    (("Skeletal Muscle Mass (kg)",), (7045, 5365, 8318, 5703)),
    (("Right Arm (kg)",), (7045, 5703, 8318, 6205)),
    (("Left Arm (kg)",), (7045, 6205, 8318, 6547)),
    (("Right Leg (kg)",), (7045, 6547, 8318, 7055)),
    (("Left Leg (kg)",), (7045, 7055, 8318, 7396)),
    (("Torso (kg)",), (7045, 7396, 8318, 7882)),
    (("Visceral Adipose Tissue",), (7045, 7882, 8318, 8218)),
    (("SECA BMI (kg/m^2)",), (7045, 8218, 8318, 8750)),
    (("Height (m)",), (7045, 8750, 8318, 9087)),
    (("Weight (kg)",), (7045, 9087, 8318, 9564)),
    (("Total Body Water (L)", "Total Body Water (%)"), (7045, 9564, 8318, 9908)),
    (("Extracellular Water (L)", "Extracellular Water (%)"), (7045, 9908, 8318, 10456)),
    (("ECW/TBW (%)",), (7045, 10456, 8318, 10788)),
    (("Resting Energy Expenditure (kcal/day)",), (7045, 10788, 8318, 11296)),
    (("Energy Consumption (kcal/day)",), (7045, 11296, 8318, 11636)),
    (("Phase Angle (deg)", "Phase Angle Percentile"), (7045, 11636, 8318, 12150)),
    (("Resistance (Ohm)",), (7045, 12150, 8318, 12486)),
    (("Reactance (Ohm)",), (7045, 12486, 8318, 12990)),
    (("Physical Activity Level",), (7045, 12990, 8318, 13325)),
]


def scale_box_to_image(box, image_size):
    """Scale a crop box from the base coordinate system to the rendered image size."""

    img_w, img_h = image_size
    sx = img_w / MEASUREMENT_BASE_WIDTH
    sy = img_h / MEASUREMENT_BASE_HEIGHT
    x0, y0, x1, y1 = box
    return (
        int(x0 * sx),
        int(y0 * sy),
        int(x1 * sx),
        int(y1 * sy),
    )

# Tell pytesseract where tesseract.exe lives
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Regex that captures floats/integers including optional comma decimal marks.
NUMBER_PATTERN = re.compile(r"-?\d+(?:[.,]\d+)?")

PATIENT_FIELDS = {
    # Use a word boundary to avoid matching the "age" portion inside other words
    # such as "Average", which previously resulted in incorrect ages (e.g. "1").
    "Age": re.compile(r"\bAge[:\s]+(\d+)", re.IGNORECASE),
}

PATIENT_METADATA_FIELDS = [
    "Patient ID",
    "Sex",
    "Age",
    "Collection Date",
    "Collection Time",
]

MEASUREMENT_FIELD_NAMES: List[str] = [
    "Fat Mass (kg)",
    "Fat Mass (%)",
    "Fat Mass Index (kg/m^2)",
    "Fat-Free Mass (kg)",
    "Fat-Free Mass (%)",
    "Fat-Free Mass Index (kg/m^2)",
    "Skeletal Muscle Mass (kg)",
    "Right Arm (kg)",
    "Left Arm (kg)",
    "Right Leg (kg)",
    "Left Leg (kg)",
    "Torso (kg)",
    "Visceral Adipose Tissue",
    "SECA BMI (kg/m^2)",
    "Height (m)",
    "Weight (kg)",
    "Total Body Water (L)",
    "Total Body Water (%)",
    "Extracellular Water (L)",
    "Extracellular Water (%)",
    "ECW/TBW (%)",
    "Resting Energy Expenditure (kcal/day)",
    "Energy Consumption (kcal/day)",
    "Phase Angle (deg)",
    "Phase Angle Percentile",
    "Resistance (Ohm)",
    "Reactance (Ohm)",
    "Physical Activity Level",
]

CALCULATED_FIELD_NAMES: List[str] = [
    "Body Mass Index (kg/m^2)",
]

REVIEWABLE_FIELDS = set(MEASUREMENT_FIELD_NAMES + CALCULATED_FIELD_NAMES)


def output_field_order() -> List[str]:
    order: List[str] = []
    for name in MEASUREMENT_FIELD_NAMES:
        order.append(name)
        if name == "SECA BMI (kg/m^2)":
            order.extend(CALCULATED_FIELD_NAMES)
    return [
        "Source File",
        *PATIENT_METADATA_FIELDS,
        "Data Quality",
        "Data Quality Fails",
        *order,
    ]


OUTPUT_FIELD_ORDER = output_field_order()

def normalize_number(token: str) -> float:
    token = token.replace(",", ".")
    return float(token)


def collapse_whitespace(text: str) -> str:
    """Return *text* with all whitespace collapsed to single spaces."""

    return " ".join(text.split())


def extract_measurements_from_page_image(
    pil_image,
) -> Tuple[Dict[str, Optional[float]], List[str], Dict[str, object]]:
    """Run OCR against each measurement crop on a rendered page image.

    Returns a tuple containing measurements, debug text, and the cropped
    ``PIL.Image`` objects for each field.
    """

    measurements: Dict[str, Optional[float]] = {}
    debug_lines: List[str] = []
    field_images: Dict[str, object] = {}

    for fields, base_box in MEASUREMENT_CROP_BOXES:
        try:
            crop_box = scale_box_to_image(base_box, pil_image.size)
            cropped = pil_image.crop(crop_box)
            ocr_text = pytesseract.image_to_string(cropped)
        except Exception:
            cropped = None
            ocr_text = ""

        cleaned_text = collapse_whitespace(ocr_text)
        if cleaned_text:
            debug_lines.append(f"{' | '.join(fields)} => {cleaned_text}")

        numbers = [normalize_number(match) for match in NUMBER_PATTERN.findall(ocr_text)]
        for index, field in enumerate(fields):
            measurements[field] = numbers[index] if index < len(numbers) else None
            if field not in field_images and cropped is not None:
                field_images[field] = cropped.copy()

    return measurements, debug_lines, field_images


def extract_measurements_from_pdf(
    pdf_path: Path,
) -> Tuple[Dict[str, Optional[float]], str, Dict[str, object]]:
    """Extract measurement values by OCR-ing each configured crop box."""

    measurements: Dict[str, Optional[float]] = {name: None for name in MEASUREMENT_FIELD_NAMES}
    debug_parts: List[str] = []
    field_images: Dict[str, object] = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            pil_image = page.to_image(resolution=OCR_RENDER_RESOLUTION).original
            (
                page_measurements,
                page_debug_lines,
                page_field_images,
            ) = extract_measurements_from_page_image(pil_image)
            for field, value in page_measurements.items():
                if value is not None:
                    measurements[field] = value
            for field, image in page_field_images.items():
                if field not in field_images:
                    field_images[field] = image

            if page_debug_lines:
                debug_parts.append(f"Page {page_index}:\n" + "\n".join(page_debug_lines))

    return measurements, "\n\n".join(debug_parts), field_images

def extract_text_layer(pdf_path: Path) -> str:
    """Extract ONLY the PDF's embedded text layer (no OCR)."""
    parts: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            parts.append(text)
    return "\n".join(parts)

def parse_patient_metadata(text: str) -> Dict[str, Optional[str]]:
    metadata: Dict[str, Optional[str]] = {
        "Patient ID": None,
        "Sex": None,
        "Age": None,
        "Collection Date": None,
        "Collection Time": None,
    }

    patient_id_match = re.search(r"ID\s*[:\-]?\s*(.*?)\s+Name", text, re.IGNORECASE)
    if patient_id_match:
        metadata["Patient ID"] = patient_id_match.group(1).strip()

    for field, pattern in PATIENT_FIELDS.items():
        match = pattern.search(text)
        if match:
            metadata[field] = match.group(1).strip()

    sex_match = re.search(r"\b(Male|Female)\b", text, re.IGNORECASE)
    if sex_match:
        metadata["Sex"] = sex_match.group(1).title()

    age_fallback = re.search(r"\b(\d{1,3})\s+(Male|Female)\b", text, re.IGNORECASE)
    if metadata["Age"] is None and age_fallback:
        metadata["Age"] = age_fallback.group(1)

    date_match = re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text)
    if date_match:
        metadata["Collection Date"] = date_match.group(1)

    time_match = re.search(r"(\d{1,2}:\d{2}\s?(?:AM|PM)?)", text, re.IGNORECASE)
    if time_match:
        metadata["Collection Time"] = time_match.group(1)

    return metadata


def evaluate_data_quality(values: Dict[str, Optional[float]]) -> Dict[str, Optional[str]]:
    def numbers_present(fields: List[str]) -> bool:
        return all(values.get(field) is not None for field in fields)

    def almost_equal(calculated: float, expected: float, tolerance: float) -> bool:
        return abs(calculated - expected) <= tolerance

    def add_decimal_between_first_two_digits(value: Optional[float]) -> Optional[float]:
        """
        Insert a decimal point after the first digit of an integer-like value.

        Examples:
            41   -> 4.1
            410  -> 4.10 -> 4.1

        Only integer-like inputs with at least two digits are adjusted; other
        values return ``None`` to signal no change.
        """

        if value is None:
            return None

        if not math.isclose(value, round(value), abs_tol=1e-6):
            return None

        sign = -1 if value < 0 else 1
        integer_part = str(int(abs(round(value))))

        if len(integer_part) < 2:
            return None

        decimal_value = float(f"{integer_part[0]}.{integer_part[1:]}")
        return sign * decimal_value

    failures: List[str] = []

    if numbers_present(["Fat Mass (kg)", "Fat-Free Mass (kg)", "Weight (kg)"]):
        fm = values["Fat Mass (kg)"]
        ffm = values["Fat-Free Mass (kg)"]
        weight = values["Weight (kg)"]
        if not almost_equal((fm or 0) + (ffm or 0), weight or 0, 0.01):
            failures.append("1")
    else:
        failures.append("1")

    if numbers_present(["Fat Mass (%)", "Fat-Free Mass (%)"]):
        fm_pct = values["Fat Mass (%)"]
        ffm_pct = values["Fat-Free Mass (%)"]
        if not almost_equal((fm_pct or 0) + (ffm_pct or 0), 100, 0.01):
            failures.append("2")
    else:
        failures.append("2")

    if numbers_present([
        "Fat Mass Index (kg/m^2)",
        "Fat-Free Mass Index (kg/m^2)",
        "SECA BMI (kg/m^2)",
    ]):
        fmi = values["Fat Mass Index (kg/m^2)"]
        ffmi = values["Fat-Free Mass Index (kg/m^2)"]
        bmi = values["SECA BMI (kg/m^2)"]
        if not almost_equal((fmi or 0) + (ffmi or 0), bmi or 0, 0.02):
            failures.append("3")
    else:
        failures.append("3")

    if numbers_present([
        "Right Arm (kg)",
        "Left Arm (kg)",
        "Right Leg (kg)",
        "Left Leg (kg)",
        "Torso (kg)",
        "Skeletal Muscle Mass (kg)",
    ]):
        sum_limbs = sum(
            values.get(field, 0) or 0
            for field in [
                "Right Arm (kg)",
                "Left Arm (kg)",
                "Right Leg (kg)",
                "Left Leg (kg)",
                "Torso (kg)",
            ]
        )
        if not almost_equal(sum_limbs, values["Skeletal Muscle Mass (kg)"] or 0, 0.03):
            failures.append("4")
    else:
        failures.append("4")

    if numbers_present(["Weight (kg)", "Height (m)", "SECA BMI (kg/m^2)"]):
        weight = values["Weight (kg)"]
        height = values["Height (m)"]
        bmi = values["SECA BMI (kg/m^2)"]
        if height in (0, None):
            failures.append("5")
        elif not almost_equal((weight or 0) / ((height or 1) ** 2), bmi or 0, 0.3):
            failures.append("5")
    else:
        failures.append("5")

    if numbers_present([
        "Extracellular Water (L)",
        "Total Body Water (L)",
        "ECW/TBW (%)",
    ]):
        ecw = values["Extracellular Water (L)"]
        tbw = values["Total Body Water (L)"]
        ratio = values["ECW/TBW (%)"]
        if tbw in (0, None):
            failures.append("6")
        elif not almost_equal(((ecw or 0) / (tbw or 1)) * 100, ratio or 0, 0.02):
            failures.append("6")
    else:
        failures.append("6")

    if numbers_present([
        "Extracellular Water (%)",
        "Total Body Water (%)",
        "ECW/TBW (%)",
    ]):
        ecw_pct = values["Extracellular Water (%)"]
        tbw_pct = values["Total Body Water (%)"]
        ratio = values["ECW/TBW (%)"]
        if tbw_pct in (0, None):
            failures.append("7")
        elif not almost_equal(((ecw_pct or 0) / (tbw_pct or 1)) * 100, ratio or 0, 0.02):
            failures.append("7")
    else:
        failures.append("7")

    if numbers_present([
        "Resting Energy Expenditure (kcal/day)",
        "Physical Activity Level",
        "Energy Consumption (kcal/day)",
    ]):
        ree = values["Resting Energy Expenditure (kcal/day)"]
        pal = values["Physical Activity Level"]
        energy = values["Energy Consumption (kcal/day)"]
        if not almost_equal((ree or 0) * (pal or 0), energy or 0, 0.02):
            failures.append("8")
    else:
        failures.append("8")

    if numbers_present(["Reactance (Ohm)", "Resistance (Ohm)", "Phase Angle (deg)"]):
        reactance = values["Reactance (Ohm)"]
        resistance = values["Resistance (Ohm)"]
        phase_angle = values["Phase Angle (deg)"]
        original_phase_angle = phase_angle
        if resistance in (0, None):
            failures.append("9")
        else:
            calculated = math.atan((reactance or 0) / (resistance or 1)) * 180 / math.pi
            if not almost_equal(calculated, phase_angle or 0, 0.1):
                adjusted_phase_angle = add_decimal_between_first_two_digits(phase_angle)

                if adjusted_phase_angle is not None and almost_equal(
                    calculated, adjusted_phase_angle, 0.1
                ):
                    values["Phase Angle (deg)"] = adjusted_phase_angle
                else:
                    values["Phase Angle (deg)"] = original_phase_angle
                    failures.append("9")
    else:
        failures.append("9")

    percentile = values.get("Phase Angle Percentile")
    if percentile is None or percentile < 0 or percentile > 100:
        failures.append("10")

    return {
        "Data Quality": "Pass" if not failures else "Fail",
        "Data Quality Fails": ",".join(failures) if failures else "",
    }


QC_FIELD_MAP = {
    "1": ["Fat Mass (kg)", "Fat-Free Mass (kg)", "Weight (kg)"],
    "2": ["Fat Mass (%)", "Fat-Free Mass (%)"],
    "3": ["Fat Mass Index (kg/m^2)", "Fat-Free Mass Index (kg/m^2)", "SECA BMI (kg/m^2)"],
    "4": [
        "Right Arm (kg)",
        "Left Arm (kg)",
        "Right Leg (kg)",
        "Left Leg (kg)",
        "Torso (kg)",
        "Skeletal Muscle Mass (kg)",
    ],
    "5": ["Weight (kg)", "Height (m)", "SECA BMI (kg/m^2)"],
    "6": ["Extracellular Water (L)", "Total Body Water (L)", "ECW/TBW (%)"],
    "7": ["Extracellular Water (%)", "Total Body Water (%)", "ECW/TBW (%)"],
    "8": [
        "Resting Energy Expenditure (kcal/day)",
        "Physical Activity Level",
        "Energy Consumption (kcal/day)",
    ],
    "9": ["Reactance (Ohm)", "Resistance (Ohm)", "Phase Angle (deg)"],
    "10": ["Phase Angle Percentile"],
}


def recompute_calculated_fields(row: Dict[str, Optional[float]]) -> None:
    weight = row.get("Weight (kg)")
    height = row.get("Height (m)")
    if height not in (None, 0):
        row["Body Mass Index (kg/m^2)"] = (
            (weight or 0) / ((height or 1) ** 2)
        ) if weight is not None else None
    else:
        row["Body Mass Index (kg/m^2)"] = None


def refresh_data_quality(row: Dict[str, Optional[float]]) -> None:
    recompute_calculated_fields(row)
    row.update(evaluate_data_quality(row))


def parse_user_value(field_name: str, value: str) -> Optional[object]:
    if field_name in MEASUREMENT_FIELD_NAMES + CALCULATED_FIELD_NAMES:
        cleaned = value.strip()
        if not cleaned:
            return None
        match = NUMBER_PATTERN.search(cleaned)
        return normalize_number(match.group(0)) if match else None
    return value if value.strip() else None


def extract_pdf_data(pdf_path: Path) -> Tuple[Dict[str, Optional[float]], Dict[str, object]]:

    row: Dict[str, Optional[float]] = {field: None for field in OUTPUT_FIELD_ORDER}
    row["Source File"] = pdf_path.name

    # --- 1. HEADER TEXT (for Patient ID, Sex, Age, Date, Time) ---
    text_layer = extract_text_layer(pdf_path)
    keyword_text = text_layer.lower()
    if "patient data" not in keyword_text or "single measurement" not in keyword_text:
        row.update(
            {
                "Data Quality": "Fail",
                "Data Quality Fails": "Not recognized as a SECA data export",
            }
        )
        return row, {}

    normalized_header_text = collapse_whitespace(text_layer)

    # --- 2. OCR TEXT (per-field cropped regions) ---
    measurements, ocr_debug_text, field_images = extract_measurements_from_pdf(pdf_path)

    # --- 3. DEBUG OUTPUT (optional) ---
    debug_txt = pdf_path.with_suffix(".ocr.txt")
    debug_txt.write_text(ocr_debug_text, encoding="utf-8")

    # --- 4. Build the row ---
    row.update(parse_patient_metadata(normalized_header_text))   # header from TEXT layer
    row.update(measurements)                                     # numbers from OCR crops

    weight = row.get("Weight (kg)")
    height = row.get("Height (m)")
    if height not in (None, 0):
        row["Body Mass Index (kg/m^2)"] = (
            (weight or 0) / ((height or 1) ** 2)
        ) if weight is not None else None

    row.update(evaluate_data_quality(row))

    return row, field_images

def select_pdf_files() -> List[Path]:
    root = Tk()
    root.withdraw()

    try:
        file_paths = filedialog.askopenfilenames(
            title="Select SECA PDF files", filetypes=[("PDF files", "*.pdf")]
        )
        root.update()
    except KeyboardInterrupt:
        # Allow users to cancel with Ctrl+C without seeing a traceback
        return []
    finally:
        root.destroy()

    return [Path(path) for path in file_paths]


def select_output_path() -> Optional[Path]:
    root = Tk()
    root.withdraw()

    try:
        directory = filedialog.askdirectory(title="Select download folder")
        root.update()
    except KeyboardInterrupt:
        # Allow users to cancel with Ctrl+C without seeing a traceback
        return None
    finally:
        root.destroy()

    if not directory:
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path(directory) / f"seca_measurements_{timestamp}.xlsx"


def show_message(title: str, message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def prompt_fix_or_continue(blank_count: int, qc_failure_count: int) -> bool:
    parts = []
    if qc_failure_count:
        parts.append(f"{qc_failure_count} quality control failure(s)")
    if blank_count:
        parts.append(f"{blank_count} blank OCR field(s)")

    summary = " and ".join(parts) if parts else "detected issues"

    root = Tk()
    root.withdraw()
    response = messagebox.askyesno(
        "Review required",
        f"Detected {summary}.\nWould you like to review and correct the data before exporting?",
        icon="warning",
    )
    root.destroy()
    return response


class PostProcessingEditor:
    def __init__(self, review_items: List[Dict[str, object]]):
        self.review_items = review_items
        self.current_index = 0
        self.decisions: Dict[Tuple[int, str], object] = {}
        self.committed = False
        self.current_photo = None

        self.root = Tk()
        self.root.title("Review OCR fields")

        self.progress_label = ttk.Label(self.root, text="")
        self.progress_label.pack(padx=10, pady=(10, 5))

        self.file_label = ttk.Label(self.root, text="")
        self.file_label.pack(padx=10, pady=5)

        self.field_label = ttk.Label(self.root, text="")
        self.field_label.pack(padx=10, pady=5)

        self.original_value_var = StringVar()
        self.original_value_label = ttk.Label(
            self.root, textvariable=self.original_value_var
        )
        self.original_value_label.pack(padx=10, pady=5)

        self.image_label = ttk.Label(self.root)
        self.image_label.pack(padx=10, pady=5)

        self.entry_var = StringVar()
        entry_label = ttk.Label(self.root, text="Edit OCR value")
        entry_label.pack(padx=10, pady=(5, 0))
        self.entry = ttk.Entry(self.root, width=50, textvariable=self.entry_var)
        self.entry.pack(padx=10, pady=(0, 10))

        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=(0, 10))
        self.back_button = ttk.Button(button_frame, text="Back", command=self.go_back)
        self.back_button.pack(side="left", padx=5)
        self.next_button = ttk.Button(button_frame, text="Next", command=self.save_and_next)
        self.next_button.pack(side="left", padx=5)
        self.save_button = ttk.Button(
            button_frame, text="Save all changes", command=self.commit_changes, state="disabled"
        )
        self.save_button.pack(side="left", padx=5)

        self.root.bind("<Left>", self.go_back)
        self.root.bind("<Right>", self.save_and_next)

        self.show_current_item()

    def format_value(self, value: Optional[object]) -> str:
        if value in (None, ""):
            return "blank"
        return str(value)

    def save_and_next(self, event=None) -> None:
        if self.current_index >= len(self.review_items):
            return

        item = self.review_items[self.current_index]
        self.decisions[(item["entry_index"], item["field"])] = self.entry_var.get()
        self.current_index += 1
        self.show_current_item()

    def go_back(self, event=None) -> None:
        if self.current_index <= 0:
            return

        self.current_index -= 1
        self.show_current_item()

    def commit_changes(self) -> None:
        self.committed = True
        self.root.destroy()

    def show_current_item(self) -> None:
        if self.current_index >= len(self.review_items):
            self.progress_label.config(text="Review complete. Click Save all changes to apply edits.")
            self.file_label.config(text="")
            self.field_label.config(text="")
            self.original_value_var.set("")
            self.entry_var.set("")
            self.entry.state(["disabled"])
            self.next_button.state(["disabled"])
            self.back_button.state(["disabled"])
            self.image_label.config(image="", text="")
            self.current_photo = None
            self.save_button.state(["!disabled"])
            return

        item = self.review_items[self.current_index]
        total = len(self.review_items)
        self.progress_label.config(
            text=f"Reviewing field {self.current_index + 1} of {total}"
        )
        self.file_label.config(text=f"File: {item['file']}")
        self.field_label.config(text=f"Field: {item['field']}")
        self.original_value_var.set(
            f"OCR detected value: {self.format_value(item.get('value'))}"
        )

        saved_value = self.decisions.get((item["entry_index"], item["field"]))
        entry_value = saved_value if saved_value is not None else item.get("value")

        if item.get("image") is not None:
            preview = item["image"].copy()
            preview.thumbnail((800, 600))
            # Bind the image explicitly to this editor's root to avoid cross-root issues
            self.current_photo = ImageTk.PhotoImage(preview, master=self.root)
            self.image_label.config(image=self.current_photo, text="")
            # Keep a reference so it's not garbage-collected
            self.image_label.image = self.current_photo
        else:
            self.image_label.config(text="No OCR snapshot available", image="")
            self.current_photo = None
            self.image_label.image = None

        self.entry.state(["!disabled"])
        self.next_button.state(["!disabled"])
        if self.current_index == 0:
            self.back_button.state(["disabled"])
        else:
            self.back_button.state(["!disabled"])
        self.save_button.state(["disabled"])
        self.entry_var.set(self.format_value(entry_value))

    def run(self) -> Tuple[Dict[Tuple[int, str], object], bool]:
        self.root.mainloop()
        return self.decisions, self.committed


def review_entries(entries: List[Dict[str, object]]) -> None:
    review_items: List[Dict[str, object]] = []
    qc_failure_count = 0
    blank_field_count = 0

    for index, entry in enumerate(entries):
        row = entry["row"]

        # Skip files that were not recognized as SECA exports; they have no OCR
        # snapshots to review and should not trigger blank-field prompts.
        if (
            row.get("Data Quality") == "Fail"
            and row.get("Data Quality Fails") == "Not recognized as a SECA data export"
        ):
            continue
        qc_codes = (
            [code for code in row.get("Data Quality Fails", "").split(",") if code]
            if row.get("Data Quality") == "Fail"
            else []
        )
        qc_fields: List[str] = []
        for code in qc_codes:
            qc_fields.extend(QC_FIELD_MAP.get(code, []))

        blank_fields = [
            field
            for field in REVIEWABLE_FIELDS
            if row.get(field) in (None, "")
        ]

        qc_failure_count += len(qc_fields)
        blank_field_count += len(blank_fields)

        fields_to_review: List[str] = []
        for field in blank_fields + qc_fields:
            if field in REVIEWABLE_FIELDS and field not in fields_to_review:
                fields_to_review.append(field)

        for field in fields_to_review:
            review_items.append(
                {
                    "entry_index": index,
                    "field": field,
                    "file": row.get("Source File", f"Entry {index + 1}"),
                    "value": row.get(field),
                    "image": entry["images"].get(field),
                }
            )

    if not review_items:
        return

    if not prompt_fix_or_continue(blank_field_count, qc_failure_count):
        return

    editor = PostProcessingEditor(review_items)
    decisions, committed = editor.run()
    if not committed:
        return

    updated_entries = set()
    for (entry_index, field), user_value in decisions.items():
        row = entries[entry_index]["row"]
        row[field] = parse_user_value(field, user_value)
        updated_entries.add(entry_index)

    for entry_index in updated_entries:
        refresh_data_quality(entries[entry_index]["row"])


class ProgressWindow:
    def __init__(self, total_files: int):
        self.root = Tk()
        self.root.title("Processing SECA files")
        self.total_files = total_files

        self.label = ttk.Label(self.root, text="Preparing to process files…")
        self.label.pack(padx=20, pady=(20, 10))

        self.progress = ttk.Progressbar(
            self.root, length=320, mode="determinate", maximum=total_files
        )
        self.progress.pack(padx=20, pady=(0, 20))

        self.root.update()

    def update_progress(self, current_index: int, filename: str) -> None:
        self.progress["value"] = current_index
        self.label.config(
            text=f"Processing {current_index}/{self.total_files}: {filename}"
        )
        self.root.update_idletasks()

    def close(self) -> None:
        self.root.destroy()


def main() -> None:
    pdf_files = select_pdf_files()
    if not pdf_files:
        show_message("SECA Data Converter", "No PDF files were selected.")
        return

    output_path = select_output_path()
    if output_path is None:
        show_message("SECA Data Converter", "No download folder was selected.")
        return

    parsed_entries: List[Dict[str, object]] = []
    progress = ProgressWindow(total_files=len(pdf_files))
    parsing_error: Optional[Tuple[Path, Exception]] = None

    for index, pdf in enumerate(pdf_files, start=1):
        try:
            row, images = extract_pdf_data(pdf)
            parsed_entries.append({"row": row, "images": images})
        except Exception as exc:  # pragma: no cover - user feedback path
            parsing_error = (pdf, exc)
            break

        progress.update_progress(index, pdf.name)

    progress.close()

    if parsing_error:
        pdf, exc = parsing_error
        show_message(
            "Parsing error",
            f"Could not parse '{pdf.name}'.\nError: {exc}",
        )
        return

    review_entries(parsed_entries)

    rows = [entry["row"] for entry in parsed_entries]
    df = pd.DataFrame(rows, columns=OUTPUT_FIELD_ORDER)

    try:
        df.to_excel(output_path, index=False)
    except ImportError as exc:  # pragma: no cover - user feedback path
        show_message(
            "Missing dependency",
            "Could not export to Excel because the 'openpyxl' package is not installed.\n"
            "Please install it and try again.\n\n"
            f"Original error: {exc}",
        )
        return
    except Exception as exc:  # pragma: no cover - user feedback path
        show_message(
            "Export failed",
            f"Could not save the Excel file.\nError: {exc}",
        )
        return

    show_message(
        "SECA Data Converter",
        f"Successfully saved data for {len(rows)} file(s) to:\n{output_path}",
    )


if __name__ == "__main__":
    main()
