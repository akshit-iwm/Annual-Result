"""
Configuration loader — reads all settings from .env file.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env from the same directory as this script
_env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(_env_path)


def _get(key: str, default: str = "") -> str:
    """Get an environment variable or return default."""
    return os.getenv(key, default).strip()


def _get_int(key: str, default: int = 0) -> int:
    """Get an environment variable as integer."""
    try:
        return int(_get(key, str(default)))
    except ValueError:
        return default

# Project root = directory containing this script
_PROJECT_ROOT = Path(__file__).resolve().parent


def _resolve_path(key: str, default: str = "") -> str:
    """Get a path from env and resolve it relative to project root."""
    raw = _get(key, default)
    if not raw:
        return raw
    p = Path(raw)
    if p.is_absolute():
        return str(p)
    return str(_PROJECT_ROOT / p)


# ── API ──────────────────────────────────────────────
API_LOGIN_URL = _get("API_LOGIN_URL")
API_UPLOAD_URL = _get("API_UPLOAD_URL")
API_DOWNLOAD_URL = _get("API_DOWNLOAD_URL")
API_USERNAME = _get("API_USERNAME")
API_PASSWORD = _get("API_PASSWORD")

# ── File Paths (resolved relative to project root) ──
CSV_FILE_PATH = _resolve_path("CSV_FILE_PATH")
TEMPLATE_FILE_PATH = _resolve_path("TEMPLATE_FILE_PATH")
OUTPUT_FOLDER_PATH = _resolve_path("OUTPUT_FOLDER_PATH")

# ── Excel Settings ───────────────────────────────────
ACCORD_CODE_SHEET = _get("ACCORD_CODE_SHEET", "Annual Raw")
ACCORD_CODE_CELL = _get("ACCORD_CODE_CELL", "A1")
REFRESH_WAIT_SECONDS = _get_int("REFRESH_WAIT_SECONDS", 30)
EXCEL_MAX_RETRIES = _get_int("EXCEL_MAX_RETRIES", 3)

# ── Upload Settings ──────────────────────────────────
UPLOAD_MAX_RETRIES = _get_int("UPLOAD_MAX_RETRIES", 3)
UPLOAD_RETRY_DELAY = _get_int("UPLOAD_RETRY_DELAY", 5)

# ── Logging ──────────────────────────────────────────
LOG_FILE_PATH = _resolve_path("LOG_FILE_PATH", "automation.log")
ERROR_LOG_FILE_PATH = _resolve_path("ERROR_LOG_FILE_PATH", "errors.log")
