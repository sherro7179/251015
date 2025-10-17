Attribute VB_Name = "modGlobals"
Option Explicit

Public Const SHEET_FILES As String = "파일"
Public Const SHEET_IO_NAMES As String = "IO_name"
Public Const SHEET_DATA_UPDATE As String = "data_update"
Public Const SHEET_SCRIPT_MOVE As String = "script_move"

Public Const FILE_TABLE_HEADER_ROW As Long = 7
Public Const FILE_TABLE_START_ROW As Long = 8

Public Const COL_FILE_NAME As Long = 1
Public Const COL_ORIGINAL_PATH As Long = 2
Public Const COL_SELECTED As Long = 3
Public Const COL_STATUS As Long = 4
Public Const COL_MESSAGE As Long = 5

Public Const BASE_PATH_CELL As String = "B2"
Public Const INCLUDE_FILTER_CELL As String = "B4"
Public Const EXCLUDE_FILTER_CELL As String = "B5"
Public Const FIND_VALUE_CELL As String = "B10"
Public Const TARGET_SHEET_CELL As String = "B12"

Public Const BACKUP_FOLDER_NAME As String = "_backup"
Public Const PROCESSED_FOLDER_NAME As String = "_processed"
Public Const LOG_FOLDER_RELATIVE As String = "vba\log"

Public Const PROGRESS_LABEL_CELL As String = "F2"
Public Const PROGRESS_PERCENT_CELL As String = "G2"
Public Const PROGRESS_STATUS_CELL As String = "H2"

Public Type SelectedFileEntry
    RowIndex As Long
    FileName As String
    OriginalPath As String
    ProcessedPath As String
End Type

Public Type TaskResult
    Success As Boolean
    Message As String
End Type

Public Enum TaskOperation
    taskUpdateFiles = 1
    taskIOChange = 2
    taskValueFind = 3
End Enum

Public gLogFilePath As String
Public gStartTick As Double
Public gTotalItems As Long
