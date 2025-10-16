from __future__ import annotations

import hashlib
import json
import random
from pathlib import Path
from typing import Iterable, List, Sequence, TypeVar

from PIL import Image

from .. import config

T = TypeVar("T")


def ensure_directories() -> None:
    config.THUMB_DIR.mkdir(parents=True, exist_ok=True)
    config.STATE_DIR.mkdir(parents=True, exist_ok=True)


def make_thumbnail(source: Path, target: Path, size: int = config.THUMB_SIZE) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)

    try:
        with Image.open(source) as image:
            image.thumbnail((size, size))
            image.convert("RGB").save(target, format="JPEG", quality=90, optimize=True)
    except Exception as exc:  # pylint: disable=broad-except
        raise RuntimeError(f"Failed to create thumbnail for {source}") from exc

    return target


def ensure_thumbnail(source: Path) -> Path:
    ensure_directories()
    relative = source.relative_to(config.DATA_DIR)
    # Deterministic hash to avoid overly nested directory structures.
    digest = hashlib.md5(str(relative).encode("utf-8"), usedforsecurity=False).hexdigest()  # noqa: S324
    filename = f"{digest}.jpg"
    thumbnail_path = config.THUMB_DIR / filename
    if not thumbnail_path.exists():
        make_thumbnail(source, thumbnail_path)
    return thumbnail_path


def stable_shuffle(items: Sequence[T], rng: random.Random) -> List[T]:
    mutable = list(items)
    rng.shuffle(mutable)
    return mutable


def save_json_atomic(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".tmp")
    with tmp_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    tmp_path.replace(path)
