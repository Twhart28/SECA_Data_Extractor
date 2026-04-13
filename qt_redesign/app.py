from __future__ import annotations

import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from PIL import Image, ImageDraw
from PySide6.QtCore import QEvent, QObject, QPointF, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QIcon, QImage, QPixmap
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QSpinBox,
    QStyledItemDelegate,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTextBrowser,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

_APP_DIR = Path(__file__).resolve().parent

try:  # noqa: E402
    from .backend import (
        DEBUG_SAVE_OCR_TXT,
        CALCULATED_FIELD_NAMES,
        OUTPUT_FIELD_ORDER,
        PATIENT_FIELDS,
        QC_FIELD_MAP,
        REVIEWABLE_FIELDS,
        center_text_cells,
        collapse_whitespace,
        evaluate_data_quality,
        extract_measurements_from_pdf,
        extract_text_layer,
        normalize_collection_date,
        normalize_collection_time,
        parse_user_value,
        refresh_data_quality,
    )
except ImportError:  # pragma: no cover - direct script execution path
    if str(_APP_DIR) not in sys.path:
        sys.path.insert(0, str(_APP_DIR))
    from backend import (
        DEBUG_SAVE_OCR_TXT,
        CALCULATED_FIELD_NAMES,
        OUTPUT_FIELD_ORDER,
        PATIENT_FIELDS,
        QC_FIELD_MAP,
        REVIEWABLE_FIELDS,
        center_text_cells,
        collapse_whitespace,
        evaluate_data_quality,
        extract_measurements_from_pdf,
        extract_text_layer,
        normalize_collection_date,
        normalize_collection_time,
        parse_user_value,
        refresh_data_quality,
    )

APP_DIR = _APP_DIR
APP_ICON = APP_DIR / "App_Logo.ico"
READ_ONLY_EDIT_FIELDS = {
    "Source File",
    "Data Quality",
    "Data Quality Fails",
    *CALCULATED_FIELD_NAMES,
}
REVIEWABLE_FIELD_ORDER = [
    field for field in OUTPUT_FIELD_ORDER if field in REVIEWABLE_FIELDS
]
LEFT_HIGHLIGHT_FIELDS = {
    "Fat Mass (kg)",
    "Fat-Free Mass (kg)",
    "Total Body Water (L)",
    "Extracellular Water (L)",
    "Phase Angle (deg)",
}
RIGHT_HIGHLIGHT_FIELDS = {
    "Fat Mass (%)",
    "Fat-Free Mass (%)",
    "Total Body Water (%)",
    "Extracellular Water (%)",
    "Phase Angle Percentile",
}
QC_CODE_DESCRIPTIONS = {
    "1": "Fat Mass (kg) + Fat-Free Mass (kg) should equal Weight (kg).",
    "2": "Fat Mass (%) + Fat-Free Mass (%) should equal 100.",
    "3": "Fat Mass Index + Fat-Free Mass Index should equal SECA BMI.",
    "4": "Arm, leg, and torso muscle values should sum to Skeletal Muscle Mass.",
    "5": "Weight and Height should reproduce SECA BMI.",
    "6": "Extracellular Water (L) divided by Total Body Water (L) should match ECW/TBW (%).",
    "7": "Extracellular Water (%) divided by Total Body Water (%) should match ECW/TBW (%).",
    "8": "Resting Energy Expenditure multiplied by Physical Activity Level should match Energy Consumption.",
    "9": "Reactance and Resistance should reproduce Phase Angle (deg).",
    "10": "Phase Angle Percentile must be between 0 and 100.",
}
REPO_URL = "https://github.com/Twhart28/SECA_Data_Converter"
CONTACT_EMAIL = "thomaswhart28@gmail.com"


def default_output_path() -> Path:
    timestamp = datetime.now().strftime("%m-%d-%y %H-%M")
    return Path.home() / "Downloads" / f"Seca Export ({timestamp}).xlsx"


def format_value(value: object) -> str:
    if value in (None, ""):
        return ""
    return str(value)


def normalize_row_precision(row: Dict[str, object]) -> None:
    bmi_value = row.get("Body Mass Index (kg/m^2)")
    if isinstance(bmi_value, (int, float)):
        row["Body Mass Index (kg/m^2)"] = round(float(bmi_value), 2)


def apply_qc6_tolerance_override(row: Dict[str, object]) -> None:
    if row.get("Data Quality Fails") == "Not recognized as a SECA data export":
        return

    ecw = row.get("Extracellular Water (L)")
    tbw = row.get("Total Body Water (L)")
    ratio = row.get("ECW/TBW (%)")

    if ecw is None or tbw in (None, 0) or ratio is None:
        return

    calculated_ratio = (float(ecw) / float(tbw)) * 100
    if abs(calculated_ratio - float(ratio)) > 0.025:
        return

    fail_codes = [
        code for code in str(row.get("Data Quality Fails", "")).split(",") if code
    ]
    if "6" not in fail_codes:
        return

    fail_codes = [code for code in fail_codes if code != "6"]
    row["Data Quality Fails"] = ",".join(fail_codes)
    row["Data Quality"] = "Pass" if not fail_codes else "Fail"


def normalized_text_token(value: Optional[str]) -> str:
    return (value or "").strip().casefold()


def extract_patient_id_from_filename_qt(pdf_path: Path) -> Optional[str]:
    stem = pdf_path.stem.strip()
    if not stem:
        return None

    first_break_index: Optional[int] = None
    for index, char in enumerate(stem):
        if char in (" ", "_"):
            first_break_index = index
            break

    if first_break_index is None:
        return stem

    next_char = stem[first_break_index + 1:first_break_index + 2]
    if next_char.upper() == "T":
        for index in range(first_break_index + 1, len(stem)):
            if stem[index] in (" ", "_"):
                return stem[:index].strip() or None
        return stem or None

    return stem[:first_break_index].strip() or None


def resolve_scanned_id(header_text: str) -> Optional[str]:
    id_name_match = re.search(r"ID\s*[:\-]?\s*(.*?)\s+Name", header_text, re.IGNORECASE)
    legacy_scanned_id = id_name_match.group(1).strip() if id_name_match else None

    name_match = re.search(r"Name:\s*(\S+)", header_text, re.IGNORECASE)
    name_token = name_match.group(1).strip() if name_match else None
    if name_token and not re.search(r"[A-Za-z]", name_token):
        name_token = None

    if legacy_scanned_id and "seca_" in legacy_scanned_id.casefold():
        return name_token or None

    if legacy_scanned_id and name_token:
        if normalized_text_token(legacy_scanned_id) == normalized_text_token(name_token):
            return name_token
        return name_token

    if name_token:
        return name_token

    return legacy_scanned_id or None


def parse_patient_metadata_qt(text: str) -> Dict[str, Optional[str]]:
    metadata: Dict[str, Optional[str]] = {
        "Scanned ID": resolve_scanned_id(text),
        "Sex": None,
        "Age": None,
        "Collection Date": None,
        "Collection Time": None,
    }

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
        metadata["Collection Date"] = normalize_collection_date(date_match.group(1))

    time_match = re.search(r"(\d{1,2}:\d{2}\s?(?:AM|PM)?)", text, re.IGNORECASE)
    if time_match:
        metadata["Collection Time"] = normalize_collection_time(time_match.group(1))

    return metadata


def extract_pdf_data_qt(
    pdf_path: Path, save_ocr_txt: bool = bool(DEBUG_SAVE_OCR_TXT)
) -> tuple[Dict[str, object], Dict[str, object]]:
    row: Dict[str, object] = {field: None for field in OUTPUT_FIELD_ORDER}
    row["Source File"] = pdf_path.name
    row["Patient ID"] = extract_patient_id_from_filename_qt(pdf_path)

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
    measurements, ocr_debug_text, field_images = extract_measurements_from_pdf(pdf_path)

    if save_ocr_txt:
        debug_txt = pdf_path.with_suffix(".ocr.txt")
        debug_txt.write_text(ocr_debug_text, encoding="utf-8")

    row.update(parse_patient_metadata_qt(normalized_header_text))
    row.update(measurements)

    weight = row.get("Weight (kg)")
    height = row.get("Height (m)")
    if height not in (None, 0):
        row["Body Mass Index (kg/m^2)"] = (
            (weight or 0) / ((height or 1) ** 2)
        ) if weight is not None else None

    row.update(evaluate_data_quality(row))
    apply_qc6_tolerance_override(row)
    return row, field_images


def pil_to_pixmap(image: Image.Image) -> QPixmap:
    rgba_image = image.convert("RGBA")
    width, height = rgba_image.size
    image_bytes = rgba_image.tobytes("raw", "RGBA")
    qimage = QImage(
        image_bytes,
        width,
        height,
        QImage.Format.Format_RGBA8888,
    ).copy()
    return QPixmap.fromImage(qimage)


def emphasized_side_for_field(field_name: str) -> Optional[str]:
    if field_name in LEFT_HIGHLIGHT_FIELDS:
        return "left"
    if field_name in RIGHT_HIGHLIGHT_FIELDS:
        return "right"
    return None


def emphasize_image_for_field(image: Image.Image, field_name: str) -> Image.Image:
    side = emphasized_side_for_field(field_name)
    if side is None:
        return image

    preview = image.convert("RGBA").copy()
    width, height = preview.size
    midpoint = width // 2

    if side == "left":
        target_box = (0, 0, midpoint - 1, height - 1)
        dim_box = (midpoint, 0, width - 1, height - 1)
    else:
        target_box = (midpoint, 0, width - 1, height - 1)
        dim_box = (0, 0, midpoint - 1, height - 1)

    overlay = Image.new("RGBA", preview.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    border_width = max(3, width // 90)

    draw.rectangle(dim_box, fill=(9, 21, 19, 120))
    draw.rectangle(target_box, fill=(15, 118, 110, 36), outline=(15, 118, 110, 255), width=border_width)

    return Image.alpha_composite(preview, overlay).convert("RGB")


def build_review_items(entries: List[Dict[str, object]]) -> List[Dict[str, object]]:
    review_items: List[Dict[str, object]] = []

    for entry_index, entry in enumerate(entries):
        row = entry["row"]
        if (
            row.get("Data Quality") == "Fail"
            and row.get("Data Quality Fails")
            == "Not recognized as a SECA data export"
        ):
            continue

        qc_codes = (
            [code for code in str(row.get("Data Quality Fails", "")).split(",") if code]
            if row.get("Data Quality") == "Fail"
            else []
        )

        qc_fields_by_name: Dict[str, List[str]] = {}
        for code in qc_codes:
            for field in QC_FIELD_MAP.get(code, []):
                qc_fields_by_name.setdefault(field, []).append(code)

        blank_fields = {
            field for field in REVIEWABLE_FIELD_ORDER if row.get(field) in (None, "")
        }

        fields_to_review = [
            field
            for field in REVIEWABLE_FIELD_ORDER
            if field in blank_fields or field in qc_fields_by_name
        ]

        for field in fields_to_review:
            reasons = []
            if field in blank_fields:
                reasons.append("Blank OCR field")
            if field in qc_fields_by_name:
                reasons.append("QC check " + ", ".join(qc_fields_by_name[field]))

            review_items.append(
                {
                    "entry_index": entry_index,
                    "field": field,
                    "file": row.get("Source File", f"Entry {entry_index + 1}"),
                    "value": row.get(field),
                    "image": entry["images"].get(field),
                    "reason": "; ".join(reasons),
                }
            )

    return review_items


def export_entries(entries: List[Dict[str, object]], output_path: Path) -> None:
    rows = [entry["row"] for entry in entries]
    for row in rows:
        normalize_row_precision(row)
    df = pd.DataFrame(rows, columns=OUTPUT_FIELD_ORDER)

    df["__sort_index"] = range(len(df))
    df["__unrecognized"] = (
        df["Data Quality Fails"] == "Not recognized as a SECA data export"
    )
    df = df.sort_values(by=["__unrecognized", "__sort_index"], kind="stable")
    df = df.drop(columns=["__sort_index", "__unrecognized"])

    df.to_excel(output_path, index=False, sheet_name="All Data")
    center_text_cells(output_path)


class ProcessingWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, pdf_paths: List[Path]):
        super().__init__()
        self.pdf_paths = pdf_paths

    def run(self) -> None:
        parsed_entries: List[Dict[str, object]] = []
        total_files = len(self.pdf_paths)

        try:
            for index, pdf_path in enumerate(self.pdf_paths, start=1):
                self.progress.emit(index - 1, total_files, f"Reading {pdf_path.name}")
                row, images = extract_pdf_data_qt(pdf_path)
                normalize_row_precision(row)
                parsed_entries.append({"row": row, "images": images, "pdf_path": pdf_path})
                self.progress.emit(index, total_files, f"Processed {pdf_path.name}")
        except Exception:
            self.failed.emit(traceback.format_exc())
            return

        self.finished.emit(parsed_entries)


class ShiftWheelHorizontalScrollFilter(QObject):
    def eventFilter(self, watched, event):
        if event.type() != QEvent.Type.Wheel:
            return super().eventFilter(watched, event)

        if not (event.modifiers() & Qt.KeyboardModifier.ShiftModifier):
            return super().eventFilter(watched, event)

        widget = watched.parentWidget() if hasattr(watched, "parentWidget") else None
        if widget is None or not hasattr(widget, "horizontalScrollBar"):
            widget = watched

        if not hasattr(widget, "horizontalScrollBar"):
            return super().eventFilter(watched, event)

        scrollbar = widget.horizontalScrollBar()
        if scrollbar is None or scrollbar.maximum() <= 0:
            return super().eventFilter(watched, event)

        delta = event.angleDelta().y() or event.angleDelta().x()
        if delta == 0:
            return super().eventFilter(watched, event)

        wheel_steps = delta / 120
        step = max(12, scrollbar.pageStep() // 6)
        scrollbar.setValue(scrollbar.value() - int(wheel_steps * step))
        event.accept()
        return True


class ImagePreview(QLabel):
    def __init__(self):
        super().__init__("No OCR snapshot selected")
        self._pixmap: Optional[QPixmap] = None
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(280, 120)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setObjectName("imagePreview")

    def set_preview(self, pixmap: Optional[QPixmap]) -> None:
        self._pixmap = pixmap
        self._update_scaled_pixmap()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._update_scaled_pixmap()

    def _update_scaled_pixmap(self) -> None:
        if self._pixmap is None or self._pixmap.isNull():
            self.setText("No OCR snapshot available")
            self.setPixmap(QPixmap())
            return

        self.setText("")
        self.setPixmap(
            self._pixmap.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        )


class ReviewValueDelegate(QStyledItemDelegate):
    advance_requested = Signal(int)

    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if hasattr(editor, "setMinimumHeight"):
            editor.setMinimumHeight(30)
        if hasattr(editor, "setStyleSheet"):
            editor.setStyleSheet("padding: 2px 6px 4px 6px;")
        if hasattr(editor, "returnPressed"):
            row = index.row()
            editor.returnPressed.connect(
                lambda row=row, editor=editor: self._commit_and_advance(row, editor)
            )
        return editor

    def _commit_and_advance(self, row: int, editor) -> None:
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QStyledItemDelegate.EndEditHint.NoHint)
        self.advance_requested.emit(row)


class PaddedLineEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if hasattr(editor, "setMinimumHeight"):
            editor.setMinimumHeight(30)
        if hasattr(editor, "setStyleSheet"):
            editor.setStyleSheet("padding: 2px 6px 4px 6px;")
        return editor


class PdfViewerDialog(QDialog):
    def __init__(self, pdf_path: Path, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.document = QPdfDocument(self)

        self.setWindowTitle(f"PDF Viewer - {pdf_path.name}")
        self.resize(980, 760)
        if APP_ICON.exists():
            self.setWindowIcon(QIcon(str(APP_ICON)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        controls = QHBoxLayout()
        self.page_label = QLabel(pdf_path.name)
        self.prev_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setPrefix("Page ")
        self.page_spin.setMinimumWidth(100)
        self.page_count_label = QLabel("of 0")
        self.zoom_out_button = QPushButton("Zoom -")
        self.fit_width_button = QPushButton("Fit width")
        self.fit_page_button = QPushButton("Fit page")
        self.zoom_in_button = QPushButton("Zoom +")

        controls.addWidget(self.page_label, 1)
        controls.addWidget(self.prev_button)
        controls.addWidget(self.next_button)
        controls.addWidget(self.page_spin)
        controls.addWidget(self.page_count_label)
        controls.addSpacing(16)
        controls.addWidget(self.zoom_out_button)
        controls.addWidget(self.fit_width_button)
        controls.addWidget(self.fit_page_button)
        controls.addWidget(self.zoom_in_button)
        layout.addLayout(controls)

        self.pdf_view = QPdfView()
        self.pdf_view.setPageMode(QPdfView.PageMode.SinglePage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self.pdf_view, 1)

        self.prev_button.clicked.connect(self.go_to_previous_page)
        self.next_button.clicked.connect(self.go_to_next_page)
        self.page_spin.valueChanged.connect(self.page_spin_changed)
        self.zoom_in_button.clicked.connect(lambda: self.adjust_zoom(1.2))
        self.zoom_out_button.clicked.connect(lambda: self.adjust_zoom(1 / 1.2))
        self.fit_width_button.clicked.connect(
            lambda: self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        )
        self.fit_page_button.clicked.connect(
            lambda: self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
        )

        navigator = self.pdf_view.pageNavigator()
        navigator.currentPageChanged.connect(self.sync_page_controls)
        navigator.backAvailableChanged.connect(self._sync_nav_buttons)
        navigator.forwardAvailableChanged.connect(self._sync_nav_buttons)

        error = self.document.load(str(pdf_path))
        if error != QPdfDocument.Error.None_:
            raise RuntimeError(f"Could not load PDF: {pdf_path.name}")

        self.pdf_view.setDocument(self.document)
        self.page_spin.setMaximum(max(1, self.document.pageCount()))
        self.page_count_label.setText(f"of {self.document.pageCount()}")
        self.sync_page_controls(0)

    def sync_page_controls(self, current_page: int) -> None:
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(current_page + 1)
        self.page_spin.blockSignals(False)
        self._sync_nav_buttons()

    def _sync_nav_buttons(self) -> None:
        current_page = self.pdf_view.pageNavigator().currentPage()
        page_count = self.document.pageCount()
        self.prev_button.setEnabled(current_page > 0)
        self.next_button.setEnabled(0 <= current_page < page_count - 1)

    def page_spin_changed(self, page_number: int) -> None:
        self.jump_to_page(page_number - 1)

    def jump_to_page(self, page_index: int) -> None:
        if page_index < 0 or page_index >= self.document.pageCount():
            return
        self.pdf_view.pageNavigator().jump(page_index, QPointF(0, 0), self.pdf_view.zoomFactor())

    def go_to_previous_page(self) -> None:
        self.jump_to_page(self.pdf_view.pageNavigator().currentPage() - 1)

    def go_to_next_page(self) -> None:
        self.jump_to_page(self.pdf_view.pageNavigator().currentPage() + 1)

    def adjust_zoom(self, multiplier: float) -> None:
        current_zoom = self.pdf_view.zoomFactor()
        if current_zoom <= 0:
            current_zoom = 1.0
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.pdf_view.setZoomFactor(max(0.25, min(current_zoom * multiplier, 5.0)))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_paths: List[Path] = []
        self.output_path: Path = default_output_path()
        self.entries: List[Dict[str, object]] = []
        self.review_items: List[Dict[str, object]] = []
        self.pending_review_edits: Dict[tuple[int, str], str] = {}
        self.pdf_viewers: List[PdfViewerDialog] = []
        self.shift_wheel_filter = ShiftWheelHorizontalScrollFilter(self)
        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessingWorker] = None
        self.updating_all_rows_table = False
        self.updating_review_table = False
        self.last_export_path: Optional[Path] = None

        self.setWindowTitle("SECA Data Converter")
        self.resize(1280, 860)
        if APP_ICON.exists():
            self.setWindowIcon(QIcon(str(APP_ICON)))

        self._build_ui()
        self._apply_styles()
        self._sync_controls()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(24, 22, 24, 22)
        root_layout.setSpacing(18)

        root_layout.addWidget(self._build_header())

        workflow = QSplitter(Qt.Orientation.Horizontal)
        workflow.setChildrenCollapsible(False)
        workflow.addWidget(self._build_setup_panel())
        workflow.addWidget(self._build_results_panel())
        workflow.setStretchFactor(0, 1)
        workflow.setStretchFactor(1, 2)
        root_layout.addWidget(workflow, 1)

        self.setCentralWidget(root)
        self._install_shift_wheel_support()

    def _build_header(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        title_row = QHBoxLayout()
        title = QLabel("SECA Data Converter")
        title.setObjectName("title")
        self.info_button = QToolButton()
        self.info_button.setObjectName("infoButton")
        self.info_button.setText("i")
        self.info_button.setToolTip("Instructions, repository link, and support contact")
        self.info_button.clicked.connect(self.show_info_dialog)

        subtitle = QLabel(
            "Select reports, process OCR, review flagged fields, and export the same Excel workbook structure."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)

        title_row.addWidget(title)
        title_row.addWidget(self.info_button, 0, Qt.AlignmentFlag.AlignVCenter)
        title_row.addStretch(1)

        layout.addLayout(title_row)
        layout.addWidget(subtitle)
        return container

    def _build_setup_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(16)

        layout.addWidget(self._section_title("1. PDF reports", "Choose one or more SECA PDF files."))

        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(210)
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout.addWidget(self.file_list, 3)

        file_buttons = QHBoxLayout()
        self.add_files_button = QPushButton("Add PDFs")
        self.remove_files_button = QPushButton("Remove selected")
        self.clear_files_button = QPushButton("Clear")
        file_buttons.addWidget(self.add_files_button)
        file_buttons.addWidget(self.remove_files_button)
        file_buttons.addWidget(self.clear_files_button)
        layout.addLayout(file_buttons)

        layout.addWidget(self._section_title("2. Export file", "Choose where the Excel workbook will be saved."))

        output_row = QHBoxLayout()
        self.output_line = QLineEdit(str(self.output_path))
        self.output_line.setPlaceholderText("Choose an .xlsx output file")
        self.browse_output_button = QPushButton("Browse")
        output_row.addWidget(self.output_line, 1)
        output_row.addWidget(self.browse_output_button)
        layout.addLayout(output_row)

        layout.addWidget(self._section_title("3. Process", "Run OCR and quality checks locally on this machine."))

        self.process_button = QPushButton("Process PDFs")
        self.process_button.setObjectName("primaryButton")
        layout.addWidget(self.process_button)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        layout.addSpacing(6)

        self.add_files_button.clicked.connect(self.add_pdf_files)
        self.remove_files_button.clicked.connect(self.remove_selected_files)
        self.clear_files_button.clicked.connect(self.clear_pdf_files)
        self.browse_output_button.clicked.connect(self.browse_output_path)
        self.output_line.textChanged.connect(self._output_path_changed)
        self.process_button.clicked.connect(self.process_files)

        return panel

    def _build_results_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        heading_row = QHBoxLayout()
        heading_row.addWidget(
            self._section_title(
                "4. Review and export",
                "Flagged fields appear first. Use the all rows tab for broader edits.",
            ),
            1,
        )
        self.export_button = QPushButton("Export Excel")
        self.export_button.setObjectName("primaryButton")
        self.open_output_button = QPushButton("Open folder")
        self.open_output_button.setEnabled(False)
        heading_row.addWidget(self.export_button)
        heading_row.addWidget(self.open_output_button)
        layout.addLayout(heading_row)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_overview_tab(), "Overview")
        self.tabs.addTab(self._build_review_tab(), "Flagged review")
        self.tabs.addTab(self._build_all_rows_tab(), "Edit all rows")
        layout.addWidget(self.tabs, 1)

        self.export_button.clicked.connect(self.export_excel)
        self.open_output_button.clicked.connect(self.open_export_folder)

        return panel

    def _build_overview_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        self.summary_label = QLabel("Process PDFs to see extraction results.")
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.overview_table = QTableWidget(0, 5)
        self.overview_table.setHorizontalHeaderLabels(
            ["Source File", "Patient ID", "Scanned ID", "Data Quality", "QC / Notes"]
        )
        self.overview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.overview_table.horizontalHeader().setStretchLastSection(True)
        self.overview_table.verticalHeader().setVisible(False)
        self.overview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.overview_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.overview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.overview_table.itemDoubleClicked.connect(self.open_selected_overview_pdf)
        self.overview_table.itemSelectionChanged.connect(self._sync_controls)
        layout.addWidget(self.overview_table, 1)

        overview_buttons = QHBoxLayout()
        self.open_selected_overview_pdf_button = QPushButton("View selected PDF")
        overview_buttons.addWidget(self.open_selected_overview_pdf_button)
        overview_buttons.addStretch(1)
        layout.addLayout(overview_buttons)

        qc_title = QLabel("QC code legend")
        qc_title.setObjectName("summaryLabel")
        layout.addWidget(qc_title)

        self.qc_legend = QTextEdit()
        self.qc_legend.setReadOnly(True)
        self.qc_legend.setMaximumHeight(150)
        self.qc_legend.setPlainText(
            "\n".join(f"{code}: {description}" for code, description in QC_CODE_DESCRIPTIONS.items())
        )
        layout.addWidget(self.qc_legend)

        self.open_selected_overview_pdf_button.clicked.connect(self.open_selected_overview_pdf)
        return tab

    def _build_review_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        self.review_status_label = QLabel("Process PDFs to populate flagged fields.")
        self.review_status_label.setObjectName("summaryLabel")
        self.review_status_label.setWordWrap(True)
        layout.addWidget(self.review_status_label)

        self.preview_title = QLabel("OCR snapshot")
        self.preview_title.setObjectName("summaryLabel")
        self.preview_title.setWordWrap(True)
        self.image_preview = ImagePreview()
        self.image_preview.setMinimumHeight(108)
        self.image_preview.setMaximumHeight(160)
        layout.addWidget(self.preview_title)
        layout.addWidget(self.image_preview)

        helper = QLabel("Type the corrected value in the highlighted cell, then press Enter to move to the next flagged field.")
        helper.setObjectName("statusLabel")
        helper.setWordWrap(True)
        layout.addWidget(helper)

        self.review_table = QTableWidget(0, 5)
        self.review_table.setHorizontalHeaderLabels(
            ["File", "Field", "Current value", "Corrected value", "Reason"]
        )
        review_header = self.review_table.horizontalHeader()
        review_header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        review_header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        review_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        review_header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        review_header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        review_header.setStretchLastSection(True)
        self.review_table.verticalHeader().setVisible(False)
        self.review_table.verticalHeader().setDefaultSectionSize(38)
        self.review_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.review_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.review_table.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        self.review_table.currentCellChanged.connect(self.review_current_cell_changed)
        self.review_table.cellClicked.connect(self.review_cell_clicked)
        self.review_delegate = ReviewValueDelegate(self.review_table)
        self.review_delegate.advance_requested.connect(self.apply_review_edit_and_advance)
        self.review_table.setItemDelegateForColumn(3, self.review_delegate)
        layout.addWidget(self.review_table, 1)

        review_buttons = QHBoxLayout()
        self.apply_selected_review_button = QPushButton("Apply selected edit")
        self.apply_all_review_button = QPushButton("Apply all review edits")
        self.open_selected_review_pdf_button = QPushButton("View source PDF")
        self.refresh_flags_button = QPushButton("Refresh flags")
        review_buttons.addWidget(self.apply_selected_review_button)
        review_buttons.addWidget(self.apply_all_review_button)
        review_buttons.addWidget(self.open_selected_review_pdf_button)
        review_buttons.addWidget(self.refresh_flags_button)
        layout.addLayout(review_buttons)

        self.apply_selected_review_button.clicked.connect(self.apply_selected_review_edit)
        self.apply_all_review_button.clicked.connect(self.apply_all_review_edits)
        self.open_selected_review_pdf_button.clicked.connect(self.open_selected_review_pdf)
        self.refresh_flags_button.clicked.connect(self.refresh_results)

        return tab

    def _build_all_rows_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        helper = QLabel(
            "Edit extracted fields directly. Quality status and calculated BMI update after cell edits."
        )
        helper.setObjectName("statusLabel")
        helper.setWordWrap(True)
        layout.addWidget(helper)

        all_rows_buttons = QHBoxLayout()
        self.open_selected_all_rows_pdf_button = QPushButton("View selected PDF")
        all_rows_buttons.addWidget(self.open_selected_all_rows_pdf_button)
        all_rows_buttons.addStretch(1)
        layout.addLayout(all_rows_buttons)

        self.all_rows_table = QTableWidget(0, len(OUTPUT_FIELD_ORDER))
        self.all_rows_table.setHorizontalHeaderLabels(OUTPUT_FIELD_ORDER)
        self.all_rows_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.all_rows_table.verticalHeader().setVisible(False)
        self.all_rows_table.verticalHeader().setDefaultSectionSize(38)
        self.all_rows_table.setAlternatingRowColors(True)
        self.all_rows_delegate = PaddedLineEditDelegate(self.all_rows_table)
        self.all_rows_table.setItemDelegate(self.all_rows_delegate)
        self.all_rows_table.itemChanged.connect(self.all_rows_item_changed)
        self.all_rows_table.itemSelectionChanged.connect(self._sync_controls)
        layout.addWidget(self.all_rows_table, 1)

        self.open_selected_all_rows_pdf_button.clicked.connect(self.open_selected_all_rows_pdf)
        return tab

    def _section_title(self, title: str, description: str) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        description_label = QLabel(description)
        description_label.setObjectName("sectionDescription")
        description_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(description_label)
        return container

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #f6f7f5;
                color: #1f2933;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QFrame#panel {
                background: #ffffff;
                border: 1px solid #d9dfdc;
                border-radius: 8px;
            }
            QLabel#title {
                font-size: 24pt;
                font-weight: 700;
                color: #17201d;
            }
            QToolButton#infoButton {
                background: #eef3f0;
                border: 1px solid #b8c5bf;
                border-radius: 12px;
                min-width: 24px;
                max-width: 24px;
                min-height: 24px;
                max-height: 24px;
                font-size: 10pt;
                font-weight: 700;
                color: #0f766e;
                padding: 0px;
            }
            QToolButton#infoButton:hover {
                background: #e2ebe6;
            }
            QLabel#subtitle,
            QLabel#sectionDescription,
            QLabel#statusLabel {
                color: #5b6770;
            }
            QLabel#sectionTitle,
            QLabel#summaryLabel {
                font-size: 11pt;
                font-weight: 650;
                color: #17201d;
            }
            QPushButton {
                background: #eef3f0;
                border: 1px solid #b8c5bf;
                border-radius: 6px;
                padding: 8px 12px;
            }
            QPushButton:hover {
                background: #e2ebe6;
            }
            QPushButton:disabled {
                color: #8b9499;
                background: #eef0ee;
                border-color: #d5d9d7;
            }
            QPushButton#primaryButton {
                background: #0f766e;
                color: #ffffff;
                border: 1px solid #0b5f59;
                font-weight: 650;
            }
            QPushButton#primaryButton:hover {
                background: #0d665f;
            }
            QListWidget,
            QLineEdit,
            QTextEdit,
            QTableWidget {
                background: #ffffff;
                border: 1px solid #cad3ce;
                border-radius: 6px;
                selection-background-color: #cce8e2;
                selection-color: #17201d;
            }
            QLineEdit {
                padding: 8px;
            }
            QTableWidget {
                gridline-color: #e2e7e4;
                alternate-background-color: #f8faf9;
            }
            QTableWidget::item:selected {
                border: 2px solid #0f766e;
                color: #17201d;
            }
            QHeaderView::section {
                background: #eef3f0;
                border: 0;
                border-right: 1px solid #d9dfdc;
                border-bottom: 1px solid #d9dfdc;
                padding: 7px;
                font-weight: 650;
            }
            QTabWidget::pane {
                border: 1px solid #d9dfdc;
                border-radius: 6px;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #eef3f0;
                border: 1px solid #d9dfdc;
                border-bottom: 0;
                padding: 8px 12px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #0f766e;
            }
            QProgressBar {
                border: 1px solid #cad3ce;
                border-radius: 6px;
                background: #ffffff;
                height: 18px;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #0f766e;
                border-radius: 5px;
            }
            QLabel#imagePreview {
                background: #f8faf9;
                border: 1px solid #d9dfdc;
                border-radius: 8px;
                color: #66727a;
            }
            """
        )

    def _install_shift_wheel_support(self) -> None:
        for widget in [
            self.file_list,
            self.overview_table,
            self.review_table,
            self.all_rows_table,
        ]:
            widget.installEventFilter(self.shift_wheel_filter)
            if hasattr(widget, "viewport"):
                widget.viewport().installEventFilter(self.shift_wheel_filter)

    def add_pdf_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select SECA PDF files",
            str(Path.home()),
            "PDF files (*.pdf)",
        )
        if not file_paths:
            return

        existing = {path.resolve() for path in self.pdf_paths}
        for file_path in file_paths:
            path = Path(file_path)
            if path.resolve() not in existing:
                self.pdf_paths.append(path)
                existing.add(path.resolve())

        self._refresh_file_list()
        self._sync_controls()

    def remove_selected_files(self) -> None:
        selected_rows = sorted(
            {index.row() for index in self.file_list.selectedIndexes()},
            reverse=True,
        )
        for row_index in selected_rows:
            del self.pdf_paths[row_index]

        self._refresh_file_list()
        self._sync_controls()

    def clear_pdf_files(self) -> None:
        self.pdf_paths.clear()
        self._refresh_file_list()
        self._sync_controls()

    def browse_output_path(self) -> None:
        selected_path, _ = QFileDialog.getSaveFileName(
            self,
            "Select Excel export location",
            str(self.output_path),
            "Excel files (*.xlsx);;All files (*.*)",
        )
        if not selected_path:
            return

        output_path = Path(selected_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
        self.output_line.setText(str(output_path))

    def _output_path_changed(self, value: str) -> None:
        self.output_path = Path(value.strip()) if value.strip() else Path()
        self._sync_controls()

    def process_files(self) -> None:
        if not self.pdf_paths:
            QMessageBox.warning(self, "No PDFs selected", "Add one or more SECA PDF files.")
            return
        if not str(self.output_path):
            QMessageBox.warning(self, "No export file", "Choose an Excel export location.")
            return

        self.entries = []
        self.review_items = []
        self.pending_review_edits = {}
        self.refresh_results()
        self.progress.setMaximum(len(self.pdf_paths))
        self.progress.setValue(0)
        self.status_label.setText("Starting OCR processing...")
        self._set_processing_state(True)

        self.thread = QThread()
        self.worker = ProcessingWorker(self.pdf_paths)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.processing_progress)
        self.worker.finished.connect(self.processing_finished)
        self.worker.failed.connect(self.processing_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self._clear_worker_refs)
        self.thread.start()

    def processing_progress(self, current: int, total: int, message: str) -> None:
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.status_label.setText(message)

    def processing_finished(self, entries: object) -> None:
        self.entries = list(entries)
        self.status_label.setText(f"Processed {len(self.entries)} file(s). Review flags before exporting.")
        self._set_processing_state(False)
        self.refresh_results()
        self.tabs.setCurrentIndex(1 if self.review_items else 0)

    def processing_failed(self, details: str) -> None:
        self.status_label.setText("Processing failed.")
        self._set_processing_state(False)
        QMessageBox.critical(
            self,
            "Processing failed",
            "The converter could not finish processing the selected PDFs.\n\n"
            f"{details}",
        )
        self.refresh_results()

    def _clear_worker_refs(self) -> None:
        self.thread = None
        self.worker = None
        self._sync_controls()

    def refresh_results(self) -> None:
        self.review_items = build_review_items(self.entries)
        valid_keys = {
            (int(item["entry_index"]), str(item["field"])) for item in self.review_items
        }
        self.pending_review_edits = {
            key: value for key, value in self.pending_review_edits.items() if key in valid_keys
        }
        self._refresh_overview()
        self._refresh_review_table()
        self._refresh_all_rows_table()
        self._sync_controls()

    def _refresh_overview(self) -> None:
        if not self.entries:
            self.summary_label.setText("Process PDFs to see extraction results.")
            self.overview_table.setRowCount(0)
            return

        total = len(self.entries)
        failed = sum(1 for entry in self.entries if entry["row"].get("Data Quality") == "Fail")
        unrecognized = sum(
            1
            for entry in self.entries
            if entry["row"].get("Data Quality Fails")
            == "Not recognized as a SECA data export"
        )
        flagged = len(self.review_items)

        self.summary_label.setText(
            f"{total} file(s) processed. {failed} failed quality checks. {flagged} field(s) need review."
        )

        self.overview_table.setRowCount(len(self.entries))
        for row_index, entry in enumerate(self.entries):
            row = entry["row"]
            notes = str(row.get("Data Quality Fails", "") or "")
            if not notes and row.get("Data Quality") == "Pass":
                notes = "Passed QC"
            values = [
                str(row.get("Source File", "") or ""),
                str(row.get("Patient ID", "") or ""),
                str(row.get("Scanned ID", "") or ""),
                str(row.get("Data Quality", "") or ""),
                notes,
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, row_index)
                self.overview_table.setItem(row_index, column_index, item)

        self.overview_table.resizeColumnsToContents()
        if self.entries:
            self.overview_table.selectRow(0)

    def _refresh_review_table(self) -> None:
        self.updating_review_table = True
        self.review_table.setRowCount(len(self.review_items))

        for row_index, item in enumerate(self.review_items):
            pending_key = (int(item["entry_index"]), str(item["field"]))
            corrected_value = self.pending_review_edits.get(
                pending_key, format_value(item.get("value"))
            )
            values = [
                item["file"],
                item["field"],
                format_value(item.get("value")),
                corrected_value,
                item["reason"],
            ]
            for column_index, value in enumerate(values):
                table_item = QTableWidgetItem(str(value))
                table_item.setData(Qt.ItemDataRole.UserRole, row_index)
                if column_index != 3:
                    table_item.setFlags(table_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if column_index == 3:
                    table_item.setToolTip("Type a corrected value and press Enter to move to the next flagged field.")
                self.review_table.setItem(row_index, column_index, table_item)

        self.review_table.resizeColumnToContents(0)
        self.review_table.resizeColumnToContents(1)
        self.review_table.resizeColumnToContents(2)
        self.review_table.resizeColumnToContents(3)
        self.review_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.review_status_label.setText(self._review_status_text())
        self.updating_review_table = False

        if self.review_items:
            self._focus_review_row(0, select_all=True)
        else:
            self.image_preview.set_preview(None)
            self.preview_title.setText("OCR snapshot")
        self._update_review_focus_styles()

    def _review_status_text(self) -> str:
        if not self.entries:
            return "Process PDFs to populate flagged fields."
        if not self.review_items:
            return "No flagged fields remain. You can still review all rows before export."
        return f"{len(self.review_items)} flagged field(s) need review."

    def _refresh_all_rows_table(self) -> None:
        self.updating_all_rows_table = True
        self.all_rows_table.setRowCount(len(self.entries))

        for row_index, entry in enumerate(self.entries):
            row = entry["row"]
            for column_index, field in enumerate(OUTPUT_FIELD_ORDER):
                table_item = QTableWidgetItem(format_value(row.get(field)))
                if field in READ_ONLY_EDIT_FIELDS:
                    table_item.setFlags(table_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table_item.setBackground(QColor("#f0f3f1"))
                self.all_rows_table.setItem(row_index, column_index, table_item)

        self.all_rows_table.resizeColumnsToContents()
        self.updating_all_rows_table = False

    def show_info_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("About SECA Data Converter")
        dialog.resize(620, 460)
        if APP_ICON.exists():
            dialog.setWindowIcon(QIcon(str(APP_ICON)))

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QLabel("SECA Data Converter")
        title.setObjectName("summaryLabel")
        body = QTextBrowser()
        body.setOpenExternalLinks(True)
        body.setHtml(
            f"""
            <h3>Instructions For New Users</h3>
            <ol>
              <li>Add one or more SECA PDF reports in the left panel.</li>
              <li>Choose the Excel export location.</li>
              <li>Click <b>Process PDFs</b> to extract data and run QC checks.</li>
              <li>Review flagged fields first. Use the OCR snapshot and PDF viewer to confirm corrections.</li>
              <li>Use <b>Edit all rows</b> if you need broader manual edits before export.</li>
              <li>Click <b>Export Excel</b> to write the workbook.</li>
            </ol>
            <h3>Notes</h3>
            <ul>
              <li>This app processes reports locally on your machine.</li>
              <li>The packaged release includes its OCR runtime. Source-checkout runs can use a local Tesseract install or SECA_TESSERACT_CMD.</li>
              <li>The Overview tab includes a QC legend explaining each QC code.</li>
            </ul>
            <h3>Links And Contact</h3>
            <p><b>GitHub repository:</b><br><a href="{REPO_URL}">{REPO_URL}</a></p>
            <p><b>Questions or issues:</b><br><a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a></p>
            """
        )

        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.accept)

        layout.addWidget(title)
        layout.addWidget(body, 1)
        layout.addWidget(close_button, 0, Qt.AlignmentFlag.AlignRight)

        dialog.exec()

    def _entry_pdf_path(self, entry_index: int) -> Optional[Path]:
        if entry_index < 0 or entry_index >= len(self.entries):
            return None

        entry = self.entries[entry_index]
        pdf_path = entry.get("pdf_path")
        if isinstance(pdf_path, Path):
            return pdf_path

        source_file = str(entry["row"].get("Source File", "") or "")
        for candidate in self.pdf_paths:
            if candidate.name == source_file:
                return candidate
        return None

    def _open_pdf_for_entry(self, entry_index: int) -> None:
        pdf_path = self._entry_pdf_path(entry_index)
        if pdf_path is None or not pdf_path.exists():
            QMessageBox.warning(self, "PDF not found", "Could not locate the source PDF for that row.")
            return
        self._show_pdf_viewer(pdf_path)

    def _show_pdf_viewer(self, pdf_path: Path) -> None:
        try:
            dialog = PdfViewerDialog(pdf_path, self)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Could not open PDF",
                f"The in-app viewer could not load this PDF.\n\n{exc}",
            )
            return

        self.pdf_viewers.append(dialog)
        dialog.destroyed.connect(lambda *_: self._forget_pdf_viewer(dialog))
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _forget_pdf_viewer(self, dialog: PdfViewerDialog) -> None:
        if dialog in self.pdf_viewers:
            self.pdf_viewers.remove(dialog)

    def open_selected_overview_pdf(self) -> None:
        row_index = self.overview_table.currentRow()
        if row_index < 0:
            return
        self._open_pdf_for_entry(row_index)

    def open_selected_review_pdf(self) -> None:
        row_index = self.review_table.currentRow()
        if row_index < 0 or row_index >= len(self.review_items):
            return
        self._open_pdf_for_entry(int(self.review_items[row_index]["entry_index"]))

    def open_selected_all_rows_pdf(self) -> None:
        row_index = self.all_rows_table.currentRow()
        if row_index < 0:
            return
        self._open_pdf_for_entry(row_index)

    def review_current_cell_changed(self, current_row: int, current_column: int, previous_row: int, previous_column: int) -> None:
        del current_column, previous_row, previous_column
        self.update_review_preview(current_row)
        self._update_review_focus_styles()

    def review_cell_clicked(self, row: int, column: int) -> None:
        if column != 3:
            self._focus_review_row(row, select_all=True)
            return
        self._focus_review_row(row, select_all=True)

    def update_review_preview(self, row_index: Optional[int] = None) -> None:
        if row_index is None:
            row_index = self.review_table.currentRow()
        if row_index is None or row_index < 0:
            self.image_preview.set_preview(None)
            self.preview_title.setText("OCR snapshot")
            return

        if row_index < 0 or row_index >= len(self.review_items):
            return

        item = self.review_items[row_index]
        side = emphasized_side_for_field(str(item["field"]))
        suffix = ""
        if side == "left":
            suffix = " | left value emphasized"
        elif side == "right":
            suffix = " | right value emphasized"
        self.preview_title.setText(f"{item['file']} | {item['field']}{suffix}")
        image = item.get("image")
        if image is None:
            self.image_preview.set_preview(None)
            return

        self.image_preview.set_preview(pil_to_pixmap(emphasize_image_for_field(image, str(item["field"]))))

    def apply_selected_review_edit(self) -> None:
        table_row = self.review_table.currentRow()
        if table_row < 0:
            return
        self._submit_review_row(table_row)

    def apply_all_review_edits(self) -> None:
        if not self.review_items:
            return

        for table_row in range(self.review_table.rowCount()):
            self._apply_review_edit_at_row(table_row)
        self.refresh_results()

    def apply_review_edit_and_advance(self, table_row: int) -> None:
        if table_row < 0 or table_row >= self.review_table.rowCount():
            return
        self._capture_pending_review_edit(table_row)
        target_row = min(table_row + 1, self.review_table.rowCount() - 1)
        self._focus_review_row(target_row, select_all=True)

    def _apply_review_edit_at_row(self, table_row: int) -> None:
        if table_row < 0 or table_row >= len(self.review_items):
            return

        item = self.review_items[table_row]
        corrected_item = self.review_table.item(table_row, 3)
        corrected_value = corrected_item.text() if corrected_item is not None else ""
        entry_index = item["entry_index"]
        field = item["field"]
        row = self.entries[entry_index]["row"]
        row[field] = parse_user_value(field, corrected_value)
        refresh_data_quality(row)
        apply_qc6_tolerance_override(row)
        normalize_row_precision(row)
        self.pending_review_edits.pop((int(entry_index), str(field)), None)

    def _capture_pending_review_edit(self, table_row: int) -> None:
        if table_row < 0 or table_row >= len(self.review_items):
            return
        item = self.review_items[table_row]
        corrected_item = self.review_table.item(table_row, 3)
        corrected_value = corrected_item.text() if corrected_item is not None else ""
        self.pending_review_edits[(int(item["entry_index"]), str(item["field"]))] = corrected_value

    def _submit_review_row(self, table_row: int) -> None:
        if table_row < 0 or table_row >= len(self.review_items):
            return

        self._apply_review_edit_at_row(table_row)
        self.refresh_results()

        if not self.review_items:
            return

        target_row = min(table_row, len(self.review_items) - 1)
        self._focus_review_row(target_row, select_all=True)

    def _focus_review_row(self, row_index: int, select_all: bool = False) -> None:
        if row_index < 0 or row_index >= self.review_table.rowCount():
            return

        self.review_table.setCurrentCell(row_index, 3)
        self.review_table.scrollToItem(self.review_table.item(row_index, 3))
        self.review_table.editItem(self.review_table.item(row_index, 3))

        if select_all:
            QTimer.singleShot(0, self._select_active_review_editor_text)

    def _select_active_review_editor_text(self) -> None:
        editor = QApplication.focusWidget()
        if hasattr(editor, "selectAll"):
            editor.selectAll()

    def _update_review_focus_styles(self) -> None:
        current_row = self.review_table.currentRow()
        for row_index in range(self.review_table.rowCount()):
            is_active_row = row_index == current_row
            for column_index in range(self.review_table.columnCount()):
                table_item = self.review_table.item(row_index, column_index)
                if table_item is None:
                    continue

                if column_index == 3:
                    color = "#ffd874" if is_active_row else "#fff2bf"
                else:
                    color = "#dff1ea" if is_active_row else ("#ffffff" if row_index % 2 == 0 else "#f8faf9")
                table_item.setBackground(QColor(color))

    def all_rows_item_changed(self, item: QTableWidgetItem) -> None:
        if self.updating_all_rows_table:
            return

        row_index = item.row()
        column_index = item.column()
        if row_index < 0 or row_index >= len(self.entries):
            return

        field = OUTPUT_FIELD_ORDER[column_index]
        if field in READ_ONLY_EDIT_FIELDS:
            return

        row = self.entries[row_index]["row"]
        row[field] = parse_user_value(field, item.text())
        refresh_data_quality(row)
        apply_qc6_tolerance_override(row)
        normalize_row_precision(row)

        self._sync_computed_cells(row_index)
        self.review_items = build_review_items(self.entries)
        self._refresh_overview()
        self._refresh_review_table()

    def _sync_computed_cells(self, row_index: int) -> None:
        self.updating_all_rows_table = True
        row = self.entries[row_index]["row"]
        for field in READ_ONLY_EDIT_FIELDS:
            if field not in OUTPUT_FIELD_ORDER:
                continue
            column_index = OUTPUT_FIELD_ORDER.index(field)
            table_item = self.all_rows_table.item(row_index, column_index)
            if table_item is None:
                table_item = QTableWidgetItem()
                self.all_rows_table.setItem(row_index, column_index, table_item)
            table_item.setText(format_value(row.get(field)))
        self.updating_all_rows_table = False

    def export_excel(self) -> None:
        if not self.entries:
            QMessageBox.warning(self, "Nothing to export", "Process PDFs before exporting.")
            return

        if self.review_items:
            response = QMessageBox.question(
                self,
                "Apply review edits?",
                "There are still flagged fields in the review tab. Apply the current review-table edits before export?",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
                | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Yes,
            )
            if response == QMessageBox.StandardButton.Cancel:
                return
            if response == QMessageBox.StandardButton.Yes:
                self.apply_all_review_edits()

        try:
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            export_entries(self.entries, self.output_path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Export failed",
                f"Could not save the Excel workbook.\n\n{exc}",
            )
            return

        self.last_export_path = self.output_path
        self.open_output_button.setEnabled(True)
        self.status_label.setText(f"Saved {len(self.entries)} row(s) to {self.output_path}")
        QMessageBox.information(
            self,
            "Export complete",
            f"Saved data for {len(self.entries)} file(s) to:\n{self.output_path}",
        )

    def open_export_folder(self) -> None:
        if not self.last_export_path:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_export_path.parent)))

    def _refresh_file_list(self) -> None:
        self.file_list.clear()
        for path in self.pdf_paths:
            self.file_list.addItem(path.name)

    def _set_processing_state(self, processing: bool) -> None:
        for widget in [
            self.add_files_button,
            self.remove_files_button,
            self.clear_files_button,
            self.browse_output_button,
            self.output_line,
            self.process_button,
            self.export_button,
        ]:
            widget.setEnabled(not processing)
        if processing:
            return
        self._sync_controls()

    def _sync_controls(self) -> None:
        has_files = bool(self.pdf_paths)
        has_output = bool(self.output_line.text().strip())
        is_processing = self.thread is not None
        has_entries = bool(self.entries)

        self.remove_files_button.setEnabled(has_files and not is_processing)
        self.clear_files_button.setEnabled(has_files and not is_processing)
        self.process_button.setEnabled(has_files and has_output and not is_processing)
        self.export_button.setEnabled(has_entries and has_output and not is_processing)
        self.apply_selected_review_button.setEnabled(bool(self.review_items))
        self.apply_all_review_button.setEnabled(bool(self.review_items))
        self.open_selected_review_pdf_button.setEnabled(bool(self.review_items) and not is_processing)
        self.open_selected_overview_pdf_button.setEnabled(
            has_entries and self.overview_table.currentRow() >= 0 and not is_processing
        )
        self.open_selected_all_rows_pdf_button.setEnabled(
            has_entries and self.all_rows_table.currentRow() >= 0 and not is_processing
        )
        self.refresh_flags_button.setEnabled(has_entries)


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
