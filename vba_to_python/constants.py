"""
Shared constants used across the VBA-to-Python conversion modules.
"""

from pathlib import Path

# Sheet names inside the control workbook
SHEET_FILES = "파일"
SHEET_IO_NAME = "IO_name"
SHEET_DATA_UPDATE = "data_update"

# Cell references within the control workbook
BASE_PATH_CELL = "B2"
INCLUDE_FILTER_CELL = "B4"
EXCLUDE_FILTER_CELL = "B5"
FIND_STRING_CELL = "B10"
TARGET_SHEET_CELL = "B12"

# Column / row offsets for the file list on the `파일` sheet
FILE_LIST_COLUMN = 1  # Column A
FILE_LIST_START_ROW = 2

# Output layout for data_update sheet (columns A-F)
DATA_UPDATE_START_ROW = 2
DATA_UPDATE_MAX_COLUMNS = 6

# Range used by the legacy IO replacement (A5:M700)
IOCHANGE_MIN_ROW = 5
IOCHANGE_MAX_ROW = 700
IOCHANGE_MIN_COL = 1   # Column A
IOCHANGE_MAX_COL = 13  # Column M

# Supported file extensions (legacy .xls is not supported by openpyxl).
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")

# Default directory to write logs if needed
DEFAULT_LOG_DIR = Path("vba") / "log"
