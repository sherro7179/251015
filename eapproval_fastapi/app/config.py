from __future__ import annotations

from functools import lru_cache
from pathlib import Path


class Settings:
    """Centralised runtime configuration."""

    def __init__(self) -> None:
        self.project_root = Path(__file__).resolve().parents[1]
        self.rules_path = self.project_root / "data" / "rules_bundle_v2.json"
        root_parent = self.project_root.parent
        self.sample_docs_dir = (
            root_parent / "docx_to" / "전자결재_샘플_결재파일_v2"
        ).resolve()
        self.regulations_dir = (
            root_parent / "docx_to" / "전자결재_샘플_규정_패키지_v2"
        ).resolve()


@lru_cache
def get_settings() -> Settings:
    """Return cached settings instance."""
    return Settings()
