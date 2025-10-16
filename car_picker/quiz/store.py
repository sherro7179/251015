from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from .. import config
from .utils import save_json_atomic


def load_history(path: Path = config.SESSION_LOG_PATH) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def append_history(entry: Dict[str, Any], path: Path = config.SESSION_LOG_PATH) -> None:
    history = load_history(path)
    history.append(entry)
    save_json_atomic(path, history)


def build_session_entry(
    score: int,
    total_questions: int,
    difficulty: str,
    duration_seconds: float,
    seed: int,
) -> Dict[str, Any]:
    return {
        "score": score,
        "total_questions": total_questions,
        "difficulty": difficulty,
        "duration_seconds": duration_seconds,
        "seed": seed,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
