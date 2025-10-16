from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from .. import config

RANDOM_SUFFIX_PATTERN = re.compile(r"_[A-Za-z0-9]{3}$")


@dataclass(frozen=True)
class CarRecord:
    """Normalized metadata for a single car image."""

    key: str
    path: str
    make: str
    model: str
    year: int
    body_style: str
    drivetrain: str
    label_easy: str
    label_medium: str
    label_hard: str

    @property
    def relative_path(self) -> Path:
        return Path(self.path)

    def to_dict(self) -> Dict[str, object]:
        return {
            "key": self.key,
            "path": self.path,
            "make": self.make,
            "model": self.model,
            "year": self.year,
            "body_style": self.body_style,
            "drivetrain": self.drivetrain,
            "label_easy": self.label_easy,
            "label_medium": self.label_medium,
            "label_hard": self.label_hard,
        }

    @classmethod
    def from_dict(cls, payload: Dict[str, object]) -> "CarRecord":
        return cls(
            key=str(payload["key"]),
            path=str(payload["path"]),
            make=str(payload["make"]),
            model=str(payload["model"]),
            year=int(payload["year"]),
            body_style=str(payload["body_style"]),
            drivetrain=str(payload["drivetrain"]),
            label_easy=str(payload["label_easy"]),
            label_medium=str(payload["label_medium"]),
            label_hard=str(payload["label_hard"]),
        )


def _normalise_token(token: str) -> str:
    return token.replace("-", " ").replace("+", " ").strip()


def parse_filename(path: Path) -> Optional[CarRecord]:
    """Parse a dataset filename into a CarRecord.

    Returns None when filename cannot be parsed or lacks mandatory fields.
    """

    stem = path.stem
    parts = stem.split("_")

    if len(parts) < 17:
        return None

    random_suffix = parts[-1]
    if not RANDOM_SUFFIX_PATTERN.match(f"_{random_suffix}"):
        # Not following expected format; attempt a naive fallback.
        base_parts = parts
    else:
        base_parts = parts[:-1]

    if len(base_parts) < 16:
        return None

    make, model, year_token = base_parts[0], base_parts[1], base_parts[2]

    if not year_token.isdigit():
        return None

    year = int(year_token)
    body_style = base_parts[15] if len(base_parts) > 15 else ""
    drivetrain = base_parts[12] if len(base_parts) > 12 else ""

    make_norm = _normalise_token(make)
    model_norm = _normalise_token(model)
    body_norm = _normalise_token(body_style)
    drivetrain_norm = _normalise_token(drivetrain)

    # Use base_parts joined (without random suffix) to detect duplicates.
    duplicate_key = "_".join(base_parts[:16])

    label_easy = make_norm
    label_medium = f"{make_norm} {model_norm}".strip()
    label_hard = f"{make_norm} {model_norm} {year}".strip()

    relative_path = path.relative_to(config.DATA_DIR)

    return CarRecord(
        key=duplicate_key,
        path=str(relative_path).replace("\\", "/"),
        make=make_norm,
        model=model_norm,
        year=year,
        body_style=body_norm,
        drivetrain=drivetrain_norm,
        label_easy=label_easy,
        label_medium=label_medium,
        label_hard=label_hard,
    )


def iter_image_files(data_dir: Path) -> Iterable[Path]:
    for extension in config.VALID_IMAGE_EXTENSIONS:
        yield from data_dir.rglob(f"*{extension}")


def build_index(
    data_dir: Path,
    dest_path: Path,
    log_fn: Optional[callable] = None,
) -> List[CarRecord]:
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    unique_records: Dict[str, CarRecord] = {}
    total_processed = 0
    total_skipped = 0

    for idx, image_path in enumerate(iter_image_files(data_dir), 1):
        total_processed += 1
        record = parse_filename(image_path)
        if record is None:
            total_skipped += 1
            continue

        if record.key in unique_records:
            # Skip duplicates with identical metadata (different random suffix).
            continue

        unique_records[record.key] = record

        if log_fn and idx % 1000 == 0:
            log_fn(f"Indexed {idx:,} files (unique: {len(unique_records):,})")

    payload = {
        "records": [record.to_dict() for record in unique_records.values()],
        "total_processed": total_processed,
        "total_unique": len(unique_records),
        "total_skipped": total_skipped,
    }

    tmp_path = dest_path.with_suffix(".tmp")
    with tmp_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    tmp_path.replace(dest_path)

    if log_fn:
        log_fn(
            f"Index build complete. Processed: {total_processed:,}, "
            f"unique: {len(unique_records):,}, skipped: {total_skipped:,}"
        )

    return list(unique_records.values())


def load_index(dest_path: Path) -> List[CarRecord]:
    if not dest_path.exists():
        raise FileNotFoundError(f"Metadata index missing at {dest_path}")

    with dest_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    records = [CarRecord.from_dict(item) for item in payload.get("records", [])]
    return records


def ensure_index(
    data_dir: Path,
    dest_path: Path,
    rebuild: bool = False,
    log_fn: Optional[callable] = None,
) -> List[CarRecord]:
    if rebuild or not dest_path.exists():
        return build_index(data_dir, dest_path, log_fn=log_fn)
    return load_index(dest_path)
