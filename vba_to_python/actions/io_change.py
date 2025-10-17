from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

from ..constants import BASE_PATH_CELL, SHEET_FILES, SHEET_IO_NAME
from ..utils import (
    extract_file_list,
    get_required_sheet,
    iter_iochange_cells,
    load_target_workbook,
    normalize_path,
    open_control_workbook,
    read_replacements,
)


def _apply_replacements(target_path: Path, replacements: List[Tuple[str, str]]) -> None:
    wb = load_target_workbook(target_path)
    try:
        if "Test Case" not in wb.sheetnames:
            raise ValueError(f"'Test Case' 시트를 찾을 수 없습니다: {target_path.name}")
        ws = wb["Test Case"]

        for row in iter_iochange_cells(ws):
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                new_value = value
                for before, after in replacements:
                    if before and before in new_value:
                        new_value = new_value.replace(before, after)
                if new_value != value:
                    cell.value = new_value

        wb.save(target_path)
    finally:
        wb.close()


def io_change(control_path: Path) -> int:
    """
    Mirror the VBA ``IOCHANGE`` macro.
    """
    with open_control_workbook(control_path) as wb:
        ws_files = get_required_sheet(wb, SHEET_FILES)
        base_raw = normalize_path(ws_files[BASE_PATH_CELL].value)
        if not base_raw:
            raise ValueError("B2 셀에 폴더 경로가 비어 있습니다.")
        base_dir = Path(base_raw).expanduser()
        if not base_dir.exists():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {base_dir}")

        filenames = extract_file_list(ws_files)
        replacements = read_replacements(get_required_sheet(wb, SHEET_IO_NAME))

    processed = 0
    for name in filenames:
        target_path = base_dir / name
        _apply_replacements(target_path, replacements)
        processed += 1

    return processed

