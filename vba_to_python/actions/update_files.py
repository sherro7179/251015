from __future__ import annotations

from pathlib import Path
from typing import List

from ..constants import BASE_PATH_CELL, SHEET_FILES
from ..utils import (
    extract_file_list,
    get_required_sheet,
    increment_suffix,
    load_target_workbook,
    normalize_path,
    open_control_workbook,
)


def _update_single_workbook(target_path: Path) -> None:
    wb = load_target_workbook(target_path)
    try:
        if "Test Case" not in wb.sheetnames:
            raise ValueError(f"'Test Case' 시트를 찾을 수 없습니다: {target_path.name}")
        ws = wb["Test Case"]

        main_name = normalize_path(ws["A2"].value)
        if not main_name:
            raise ValueError(f"A2 셀의 기준 ID가 비어 있습니다: {target_path.name}")

        depth1 = f"{main_name}_00"
        depth2 = f"{depth1}_01"
        ws["A3"].value = depth1
        ws["A4"].value = depth2

        row = 5
        while True:
            cell = ws.cell(row=row, column=1)
            current = normalize_path(cell.value)
            if not current:
                break

            if len(current) == len(depth1):
                depth1 = increment_suffix(depth1)
                cell.value = depth1
                depth2 = f"{depth1}_00"
            elif len(current) == len(depth2):
                descriptor = normalize_path(ws.cell(row=row, column=2).value) or ""
                if "precondition" in descriptor.lower():
                    cell.value = depth2
                else:
                    depth2 = increment_suffix(depth2)
                    cell.value = depth2
            else:
                raise ValueError(
                    f"예상하지 못한 ID 패턴입니다 (행 {row}): {current}"
                )
            row += 1

        wb.save(target_path)
    finally:
        wb.close()


def update_files(control_path: Path) -> int:
    """
    Mirror the VBA ``UpdateFiles`` macro.

    Parameters
    ----------
    control_path:
        Path to the control workbook containing the ``파일`` sheet.

    Returns
    -------
    int
        Number of workbooks successfully updated.
    """
    with open_control_workbook(control_path) as wb:
        ws_files = get_required_sheet(wb, SHEET_FILES)
        base_raw = normalize_path(ws_files[BASE_PATH_CELL].value)
        if not base_raw:
            raise ValueError("B2 셀에 폴더 경로가 비어 있습니다.")

        base_dir = Path(base_raw).expanduser()
        if not base_dir.exists():
            raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {base_dir}")

        filenames: List[str] = extract_file_list(ws_files)

    processed = 0
    for name in filenames:
        target_path = base_dir / name
        _update_single_workbook(target_path)
        processed += 1

    return processed

