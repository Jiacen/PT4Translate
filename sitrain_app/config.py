from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = Path(os.getenv("SITRAIN_DATA_ROOT", ROOT / "sitrain_data")).resolve()
DB_PATH = DATA_ROOT / "sitrain.db"
STORAGE_ROOT = DATA_ROOT / "storage"
MAX_UPLOAD_BYTES = int(os.getenv("SITRAIN_MAX_UPLOAD_BYTES", 512 * 1024 * 1024))
MAX_ARCHIVE_ENTRIES = 20_000
MAX_ARCHIVE_UNCOMPRESSED = 2 * 1024 * 1024 * 1024

ALLOWED_EXTENSIONS = {
    ".pptx", ".ppt", ".docx", ".doc", ".xlsx", ".xls", ".pdf",
    ".csv", ".tsv", ".txt", ".md", ".json", ".html", ".htm",
}
LEGACY_OFFICE = {".ppt": ".pptx", ".doc": ".docx", ".xls": ".xlsx"}
MODERN_OFFICE = {".pptx", ".docx", ".xlsx"}


def ensure_data_dirs() -> None:
    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    STORAGE_ROOT.mkdir(parents=True, exist_ok=True)
