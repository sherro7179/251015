from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

from ..constants import (
    BASE_PATH_CELL,
    DATA_UPDATE_START_ROW,
    FIND_STRING_CELL,
    SHEET_DATA_UPDATE,
    SHEET_FILES,
    TARGET_SHEET_CELL,
)
from ..utils import (
    clear_data_update,
    extract_file_list,
    get_required_sheet,
    load_target_workbook,
    normalize_path,
    open_control_workbook,
)


def _scan_workbook(
    target_path: Path,
    needle: str,
) -> List[Tuple[str, str, str]]:
    """
    Return a list of matches found inside the Test Case sheet.

    Each tuple contains:
        (matched_value, neighbor_value, neighbor_address)
    """
    matches: List[Tuple[str, str, str]] = []
    wb = load_target_workbook(target_path)
    try:
        if "Test Case" not in wb.sheetnames:
            raise ValueError(f"'Test Case' 시트를 찾을 수 없습니다: {target_path.name}")
        ws = wb["Test Case"]

        needle_lower = needle.lower()
        for row in ws.iter_rows(min_row=1, min_col=3, max_col=6):
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                if needle_lower not in value.lower():
                    continue
                right_cell = ws.cell(row=cell.row, column=cell.column + 1)
                matches.append(
                    (
                        value,
                        right_cell.value,
                        right_cell.coordinate,
                    )
                )

        return matches
    finally:
        wb.close()


def value_find(control_path: Path) -> int:
    """
    Mirror the VBA ``Value_find`` macro.

    Returns the number of matches written to ``data_update``.
    """
    with open_control_workbook(control_path) as wb:
        ws_files = get_required_sheet(wb, SHEET_FILES)
        base_raw = normalize_path(ws_files[BASE_PATH_CELL].value)
        if not base_raw:
            raise ValueError("B2 셀에 폴더 경로가 비어 있습니다.")
        base_dir = Path(base_raw).expanduser()
        if not base_dir.exists():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {base_dir}")

        find_text = normalize_path(ws_files[FIND_STRING_CELL].value)
        target_sheet = normalize_path(ws_files[TARGET_SHEET_CELL].value)
        if not find_text:
            raise ValueError("B10 셀에 찾을 문자 값을 입력해 주세요.")
        if not target_sheet:
            raise ValueError("B12 셀에 대상 시트명을 입력해 주세요.")

        filenames = extract_file_list(ws_files)

        ws_update = get_required_sheet(wb, SHEET_DATA_UPDATE)
        clear_data_update(ws_update)
        next_row = DATA_UPDATE_START_ROW

        total_matches = 0
        for name in filenames:
            target_path = base_dir / name
            hits = _scan_workbook(target_path, find_text)
            for match_value, neighbor_value, neighbor_address in hits:
                ws_update.cell(row=next_row, column=1).value = str(target_path)
                ws_update.cell(row=next_row, column=2).value = match_value
                ws_update.cell(row=next_row, column=3).value = target_sheet
                ws_update.cell(row=next_row, column=4).value = neighbor_value
                ws_update.cell(row=next_row, column=5).value = neighbor_address
                next_row += 1
                total_matches += 1

        wb.save(control_path)

    return total_matches
