from __future__ import annotations

from pathlib import Path
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

from ..constants import BASE_PATH_CELL, FILE_LIST_COLUMN, FILE_LIST_START_ROW, SHEET_FILES
from ..utils import (
    clear_column,
    get_required_sheet,
    list_excel_filenames,
    normalize_path,
    open_control_workbook,
)


def _populate_file_column(ws_files: Worksheet, folder: Path) -> int:
    clear_column(ws_files, FILE_LIST_COLUMN, FILE_LIST_START_ROW)
    files = list_excel_filenames(folder)
    row = FILE_LIST_START_ROW
    for name in files:
        ws_files.cell(row=row, column=FILE_LIST_COLUMN).value = name
        row += 1
    return len(files)


def list_excel_files(control_path: Path, folder_override: Optional[Path] = None) -> int:
    """
    Mirror the VBA ``ListExcelFilesInFolder`` macro.

    Parameters
    ----------
    control_path:
        Path to the control workbook containing the ``파일`` sheet.
    folder_override:
        Optional folder to scan. When omitted, the path stored in B2 is used.

    Returns
    -------
    int
        Number of Excel files enumerated.
    """
    with open_control_workbook(control_path) as wb:
        ws_files = get_required_sheet(wb, SHEET_FILES)

        raw_path = folder_override or normalize_path(ws_files[BASE_PATH_CELL].value)
        if not raw_path:
            raise ValueError("B2 셀에 유효한 폴더 경로가 없습니다.")

        folder = Path(raw_path).expanduser()
        if not folder.exists() or not folder.is_dir():
            raise NotADirectoryError(f"폴더를 찾을 수 없습니다: {folder}")

        count = _populate_file_column(ws_files, folder)
        wb.save(control_path)

    return count

