# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path


SPEC_DIR = Path("qt_redesign").resolve()
TESSERACT_DIR = Path(
    os.environ.get("SECA_TESSERACT_DIR", r"C:\Program Files\Tesseract-OCR")
).resolve()


def collect_tesseract_runtime(source_dir: Path):
    if not source_dir.exists():
        raise FileNotFoundError(
            f"Tesseract runtime was not found at {source_dir}. "
            "Set SECA_TESSERACT_DIR to the local Tesseract install before building."
        )

    entries = [(str(SPEC_DIR / "App_Logo.ico"), ".")]

    for pattern in ("tesseract.exe", "*.dll"):
        for file_path in sorted(source_dir.glob(pattern)):
            entries.append((str(file_path), "tesseract"))

    tessdata_dir = source_dir / "tessdata"
    for subdir_name in ("configs", "script", "tessconfigs"):
        subdir = tessdata_dir / subdir_name
        if subdir.exists():
            for file_path in sorted(subdir.rglob("*")):
                if file_path.is_file():
                    dest_dir = str(Path("tesseract") / "tessdata" / subdir_name / file_path.relative_to(subdir).parent)
                    entries.append((str(file_path), dest_dir))

    for file_name in ("eng.traineddata", "osd.traineddata", "pdf.ttf"):
        file_path = tessdata_dir / file_name
        if file_path.exists():
            entries.append((str(file_path), str(Path("tesseract") / "tessdata")))

    return entries


a = Analysis(
    [str(SPEC_DIR / "app.py")],
    pathex=[str(SPEC_DIR)],
    binaries=[],
    datas=collect_tesseract_runtime(TESSERACT_DIR),
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="seca_data_extractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(SPEC_DIR / "App_Logo.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="seca_data_extractor",
)
