"""
Build and verify the Zenodo Version 1.0 upload package.
"""

import hashlib
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


PROJECT_ROOT = Path(__file__).resolve().parent.parent
PACKAGE_FILE = (
    PROJECT_ROOT / "Study_Hours_Marks_Prediction_Zenodo_Package_v1.0.zip"
)
CHECKSUM_FILE = PROJECT_ROOT / "CHECKSUMS.sha256"

FILES_TO_PACKAGE = [
    "Study_Hours_Marks_Prediction_IEEE_Preprint_v1.0.pdf",
    "Study_Hours_Marks_Prediction_IEEE_Research_Paper.docx",
    "study_hours_marks_dataset.csv",
    "main.py",
    "research_evaluation.py",
    "test_main.py",
    "requirements.txt",
    "README.md",
    "LICENSE.md",
    "SOFTWARE_LICENSE.md",
    "CITATION.cff",
    "zenodo_metadata.json",
    "ZENODO_PUBLISHING_GUIDE.md",
    "images/study_hours_marks_graph.png",
    "tools/build_ieee_research_paper.py",
    "tools/build_zenodo_pdf.py",
]


def calculate_sha256(file_path):
    """Return the SHA-256 checksum for one file."""

    checksum = hashlib.sha256()

    with file_path.open("rb") as file_handle:
        for block in iter(lambda: file_handle.read(65536), b""):
            checksum.update(block)

    return checksum.hexdigest()


def build_checksum_file():
    """Create a checksum list for all publication files."""

    lines = []

    for relative_name in FILES_TO_PACKAGE:
        file_path = PROJECT_ROOT / relative_name

        if not file_path.exists():
            raise FileNotFoundError(f"Missing package file: {relative_name}")

        lines.append(f"{calculate_sha256(file_path)}  {relative_name}")

    CHECKSUM_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_zip_package():
    """Create the final ZIP and confirm that all expected files are present."""

    build_checksum_file()
    package_files = FILES_TO_PACKAGE + [CHECKSUM_FILE.name]

    with ZipFile(PACKAGE_FILE, "w", compression=ZIP_DEFLATED) as zip_file:
        for relative_name in package_files:
            zip_file.write(PROJECT_ROOT / relative_name, arcname=relative_name)

    with ZipFile(PACKAGE_FILE, "r") as zip_file:
        bad_file = zip_file.testzip()

        if bad_file:
            raise ValueError(f"ZIP file failed verification: {bad_file}")

        actual_files = set(zip_file.namelist())
        expected_files = set(package_files)

        if actual_files != expected_files:
            raise ValueError("ZIP file list does not match the expected package.")

    print(PACKAGE_FILE)
    print(f"files={len(package_files)}")
    print(f"sha256={calculate_sha256(PACKAGE_FILE)}")


if __name__ == "__main__":
    build_zip_package()
