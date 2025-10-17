from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

from ..constants import DATA_UPDATE_START_ROW, SHEET_DATA_UPDATE
from ..utils import (
    get_required_sheet,
    is_valid_cell_address,
    load_target_workbook,
    normalize_path,
    open_control_workbook,
)


def change_value(control_path: Path) -> Tuple[int, int, List[str]]:
    """
    Mirror the VBA ``Change_Value`` macro.

    Returns a tuple of (success_count, failure_count, error_messages).
    """
    tasks: List[Tuple[Path, str, str, object]] = []

    with open_control_workbook(control_path) as wb:
        ws_update = get_required_sheet(wb, SHEET_DATA_UPDATE)
        last_row = ws_update.max_row
        if last_row < DATA_UPDATE_START_ROW:
            return (0, 0, [])

        for row in range(DATA_UPDATE_START_ROW, last_row + 1):
            raw_path = normalize_path(ws_update.cell(row=row, column=1).value)
            target_sheet = normalize_path(ws_update.cell(row=row, column=3).value)
            cell_address = normalize_path(ws_update.cell(row=row, column=5).value)
            new_value = ws_update.cell(row=row, column=6).value

            if not raw_path or not target_sheet or not cell_address:
                continue
            tasks.append((Path(raw_path).expanduser(), target_sheet, cell_address, new_value))

    success = 0
    failure = 0
    errors: List[str] = []

    for file_path, sheet_name, cell_address, new_value in tasks:
        if not file_path.exists():
            failure += 1
            errors.append(f"{file_path} - 파일을 찾을 수 없습니다.")
            continue
        if not is_valid_cell_address(cell_address):
            failure += 1
            errors.append(f"{file_path} - 잘못된 셀 주소: {cell_address}")
            continue

        try:
            wb = load_target_workbook(file_path)
            try:
                if sheet_name not in wb.sheetnames:
                    raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")
                ws = wb[sheet_name]
                ws[cell_address].value = new_value
                wb.save(file_path)
                success += 1
            finally:
                wb.close()
        except Exception as exc:
            failure += 1
            errors.append(f"{file_path} - {exc}")

    return (success, failure, errors)

