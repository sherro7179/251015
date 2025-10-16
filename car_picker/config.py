from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
METADATA_PATH = BASE_DIR / "metadata" / "index.json"
THUMB_DIR = BASE_DIR / "static" / "thumbs"
STATE_DIR = BASE_DIR / "_state"
SESSION_LOG_PATH = STATE_DIR / "sessions.json"

THUMB_SIZE = 512
DEFAULT_SESSION_LENGTH = 20
DEFAULT_DIFFICULTY = "hard"

DIFFICULTY_LEVELS = {
    "easy": "Make only",
    "medium": "Make + model",
    "hard": "Make + model + year",
}

VALID_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
