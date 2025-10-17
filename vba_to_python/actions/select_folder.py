from __future__ import annotations

from pathlib import Path

from ..constants import BASE_PATH_CELL, SHEET_FILES
from ..utils import ensure_trailing_sep, get_required_sheet, open_control_workbook
from .list_files import _populate_file_column


def select_folder_path(control_path: Path, folder_path: Path) -> int:
    """
    Mirror the VBA ``SelectFolderPath`` macro: store the folder path in B2 and
    immediately refresh the Excel file list.

    Parameters
    ----------
    control_path:
        Path to the control workbook containing the ``파일`` sheet.
    folder_path:
        Folder chosen by the user.

    Returns
    -------
    int
        Number of Excel files discovered in the selected folder.
    """
    folder = Path(folder_path).expanduser()
    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"폴더를 찾을 수 없습니다: {folder}")

    with open_control_workbook(control_path) as wb:
        ws_files = get_required_sheet(wb, SHEET_FILES)
        ws_files[BASE_PATH_CELL].value = ensure_trailing_sep(str(folder.resolve()))

        count = _populate_file_column(ws_files, folder)
        wb.save(control_path)

    return count

