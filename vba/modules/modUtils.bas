Attribute VB_Name = "modUtils"
Option Explicit

Private Const STATUS_READY As String = "Ready"

Public Sub EnsureInfrastructure()
    Dim required As Variant
    Dim nameVal As Variant
    required = Array(SHEET_FILES, SHEET_IO_NAMES, SHEET_DATA_UPDATE, SHEET_SCRIPT_MOVE)
    For Each nameVal In required
        If Not SheetExists(CStr(nameVal)) Then
            Err.Raise vbObjectError + 100, "EnsureInfrastructure", "시트 '" & nameVal & "' 을(를) 찾을 수 없습니다."
        End If
    Next nameVal
    InitializeFilesSheet
    InitializeDataUpdateSheet
End Sub

Public Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Public Function WorkbookRoot() As String
    WorkbookRoot = ThisWorkbook.Path
End Function

Public Function CombinePath(ByVal basePath As String, ByVal tailPath As String) As String
    If Right$(basePath, 1) = "\" Then
        CombinePath = basePath & tailPath
    Else
        CombinePath = basePath & "\" & tailPath
    End If
End Function

Public Function EnsureTrailingSlash(ByVal folderPath As String) As String
    If folderPath = vbNullString Then
        EnsureTrailingSlash = ""
    ElseIf Right$(folderPath, 1) = "\" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Public Sub CreateFolderIfMissing(ByVal folderPath As String)
    Dim parts() As String
    Dim idx As Long
    Dim current As String

    If Len(folderPath) = 0 Then Exit Sub
    folderPath = Replace(folderPath, "/", "\")
    parts = Split(folderPath, "\")
    current = ""
    For idx = LBound(parts) To UBound(parts)
        If parts(idx) <> "" Then
            current = current & parts(idx) & "\"
            If Dir(current, vbDirectory) = vbNullString Then
                MkDir current
            End If
        End If
    Next idx
End Sub

Public Function GetLogFolder() As String
    Dim folderPath As String
    folderPath = CombinePath(EnsureTrailingSlash(WorkbookRoot), LOG_FOLDER_RELATIVE)
    CreateFolderIfMissing folderPath
    GetLogFolder = folderPath
End Function

Public Sub InitializeFilesSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    ws.Range(BASE_PATH_CELL).EntireRow.Hidden = False
    ws.Range(BASE_PATH_CELL).EntireRow.RowHeight = ws.StandardHeight

    If ws.Cells(FILE_TABLE_HEADER_ROW, COL_FILE_NAME).Value <> "File Name" Then
        ws.Cells(FILE_TABLE_HEADER_ROW, COL_FILE_NAME).Resize(1, 5).Value = Array("File Name", "Original Path", "Include?", "Status", "Message")
        ws.Columns(COL_SELECTED).ColumnWidth = 10
        ws.Columns(COL_STATUS).ColumnWidth = 16
        ws.Columns(COL_MESSAGE).ColumnWidth = 45
        ws.Range(INCLUDE_FILTER_CELL).Value = ""
        ws.Range(EXCLUDE_FILTER_CELL).Value = ""
        ws.Range(PROGRESS_LABEL_CELL).Value = STATUS_READY
        ws.Range(PROGRESS_PERCENT_CELL).Value = 0
        ws.Range(PROGRESS_STATUS_CELL).Value = ""
    End If
End Sub

Public Sub InitializeDataUpdateSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DATA_UPDATE)
    If ws.Cells(1, 1).Value = vbNullString Then
        ws.Range("A1:F1").Value = Array("File Path", "Match Value", "Target Sheet", "Adjacent Value", "Cell Address", "New Value")
    End If
End Sub

Public Sub ClearFileTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    ws.Rows(FILE_TABLE_START_ROW & ":" & ws.Rows.Count).EntireRow.ClearContents
    ws.Range(PROGRESS_LABEL_CELL).Value = STATUS_READY
    ws.Range(PROGRESS_PERCENT_CELL).Value = 0
    ws.Range(PROGRESS_STATUS_CELL).Value = ""
End Sub

Public Function ListExcelFilesInFolder(ByVal folderPath As String) As Long
    Dim ws As Worksheet
    Dim fileName As String
    Dim rowIndex As Long
    Dim countFiles As Long
    Dim targetPath As String

    EnsureInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    ClearFileTable
    ClearStatusColors

    folderPath = EnsureTrailingSlash(folderPath)
    If Len(folderPath) = 0 Then
        Err.Raise vbObjectError + 101, "ListExcelFilesInFolder", "폴더 경로가 비어 있습니다."
    End If

    fileName = Dir(folderPath & "*.xls*")

    rowIndex = FILE_TABLE_START_ROW
    Do While Len(fileName) > 0
        ws.Cells(rowIndex, COL_FILE_NAME).Value = fileName
        targetPath = folderPath & fileName
        ws.Cells(rowIndex, COL_ORIGINAL_PATH).Value = targetPath
        ws.Cells(rowIndex, COL_SELECTED).Value = True
        ws.Cells(rowIndex, COL_STATUS).ClearContents
        ws.Cells(rowIndex, COL_STATUS).Interior.ColorIndex = xlColorIndexNone
        ws.Cells(rowIndex, COL_MESSAGE).ClearContents
        countFiles = countFiles + 1
        rowIndex = rowIndex + 1
        fileName = Dir
    Loop

    RefreshSelectionByFilters
    ListExcelFilesInFolder = countFiles
End Function

Private Function GetFilterTokens(ByVal csvText As String) As Variant
    Dim cleaned As String
    Dim tokens() As String
    Dim idx As Long

    cleaned = Trim$(csvText)
    If cleaned = vbNullString Then
        GetFilterTokens = Empty
        Exit Function
    End If

    tokens = Split(cleaned, ";")
    For idx = LBound(tokens) To UBound(tokens)
        tokens(idx) = LCase$(Trim$(tokens(idx)))
    Next idx
    GetFilterTokens = tokens
End Function

Private Function MatchesFilterTokens(ByVal valueText As String, ByVal tokens As Variant, ByVal defaultWhenEmpty As Boolean) As Boolean
    Dim idx As Long
    Dim token As String

    If IsEmpty(tokens) Then
        MatchesFilterTokens = defaultWhenEmpty
        Exit Function
    End If
    If Not IsArray(tokens) Then
        MatchesFilterTokens = defaultWhenEmpty
        Exit Function
    End If

    valueText = LCase$(valueText)

    For idx = LBound(tokens) To UBound(tokens)
        token = tokens(idx)
        If token <> vbNullString And InStr(valueText, token) > 0 Then
            MatchesFilterTokens = True
            Exit Function
        End If
    Next idx

    MatchesFilterTokens = False
End Function

Public Sub RefreshSelectionByFilters()
    Dim ws As Worksheet
    Dim includeTokens As Variant
    Dim excludeTokens As Variant
    Dim rowIndex As Long
    Dim fileName As String
    Dim manualValue As Variant
    Dim autoSelect As Boolean

    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    includeTokens = GetFilterTokens(ws.Range(INCLUDE_FILTER_CELL).Value)
    excludeTokens = GetFilterTokens(ws.Range(EXCLUDE_FILTER_CELL).Value)

    rowIndex = FILE_TABLE_START_ROW
    Do While Len(ws.Cells(rowIndex, COL_FILE_NAME).Value) > 0
        fileName = ws.Cells(rowIndex, COL_FILE_NAME).Value
        manualValue = ws.Cells(rowIndex, COL_SELECTED).Value
        autoSelect = True
        If Not IsEmpty(includeTokens) Then
            autoSelect = MatchesFilterTokens(fileName, includeTokens, True)
        End If
        If autoSelect And Not IsEmpty(excludeTokens) Then
            If MatchesFilterTokens(fileName, excludeTokens, False) Then
                autoSelect = False
            End If
        End If

        If IsEmpty(manualValue) Or manualValue = vbNullString Then
            ws.Cells(rowIndex, COL_SELECTED).Value = autoSelect
        End If
        rowIndex = rowIndex + 1
    Loop
End Sub

Public Function GetSelectedFiles(ByVal basePath As String, ByRef entries() As SelectedFileEntry) As Long
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim countSelected As Long
    Dim includeTokens As Variant
    Dim excludeTokens As Variant
    Dim fileName As String
    Dim manualValue As Variant
    Dim autoSelect As Boolean
    Dim originalPath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    includeTokens = GetFilterTokens(ws.Range(INCLUDE_FILTER_CELL).Value)
    excludeTokens = GetFilterTokens(ws.Range(EXCLUDE_FILTER_CELL).Value)

    rowIndex = FILE_TABLE_START_ROW
    Do While Len(ws.Cells(rowIndex, COL_FILE_NAME).Value) > 0
        fileName = ws.Cells(rowIndex, COL_FILE_NAME).Value
        manualValue = ws.Cells(rowIndex, COL_SELECTED).Value
        autoSelect = True
        If Not IsEmpty(includeTokens) Then
            autoSelect = MatchesFilterTokens(fileName, includeTokens, True)
        End If
        If autoSelect And Not IsEmpty(excludeTokens) Then
            If MatchesFilterTokens(fileName, excludeTokens, False) Then
                autoSelect = False
            End If
        End If

        If IsEmpty(manualValue) Or manualValue = vbNullString Then
            manualValue = autoSelect
            ws.Cells(rowIndex, COL_SELECTED).Value = manualValue
        End If

        If CBool(manualValue) Then
            countSelected = countSelected + 1
            ReDim Preserve entries(1 To countSelected)
            entries(countSelected).RowIndex = rowIndex
            entries(countSelected).FileName = fileName
            originalPath = ws.Cells(rowIndex, COL_ORIGINAL_PATH).Value
            If Len(originalPath) = 0 Then
                originalPath = EnsureTrailingSlash(basePath) & fileName
            End If
            entries(countSelected).OriginalPath = originalPath
        End If

        rowIndex = rowIndex + 1
    Loop

    GetSelectedFiles = countSelected
End Function

Public Function PrepareProcessedFile(ByVal originalPath As String) As String
    Dim baseFolder As String
    Dim fileName As String
    Dim backupFolder As String
    Dim processedFolder As String
    Dim backupPath As String
    Dim processedPath As String

    If Len(originalPath) = 0 Then
        Err.Raise vbObjectError + 102, "PrepareProcessedFile", "원본 파일 경로가 비어 있습니다."
    End If
    If Dir(originalPath, vbNormal) = vbNullString Then
        Err.Raise vbObjectError + 103, "PrepareProcessedFile", "파일을 찾을 수 없습니다: " & originalPath
    End If

    baseFolder = Left$(originalPath, InStrRev(originalPath, "\"))
    fileName = Mid$(originalPath, InStrRev(originalPath, "\") + 1)

    backupFolder = EnsureTrailingSlash(baseFolder & BACKUP_FOLDER_NAME)
    processedFolder = EnsureTrailingSlash(baseFolder & PROCESSED_FOLDER_NAME)

    CreateFolderIfMissing backupFolder
    CreateFolderIfMissing processedFolder

    backupPath = backupFolder & fileName
    processedPath = processedFolder & fileName

    If Dir(backupPath, vbNormal) = vbNullString Then FileCopy originalPath, backupPath
    If Dir(processedPath, vbNormal) <> vbNullString Then Kill processedPath
    FileCopy originalPath, processedPath

    PrepareProcessedFile = processedPath
End Function

Public Sub StartProgress(ByVal totalItems As Long, ByVal taskLabel As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    gTotalItems = totalItems
    gStartTick = Timer
    ws.Range(PROGRESS_LABEL_CELL).Value = taskLabel
    ws.Range(PROGRESS_PERCENT_CELL).Value = 0
    ws.Range(PROGRESS_STATUS_CELL).Value = "0 / " & totalItems
    Application.StatusBar = taskLabel & " 준비 중..."
End Sub

Public Sub UpdateProgress(ByVal currentIndex As Long, ByVal currentFile As String)
    Dim percentComplete As Double
    Dim elapsed As Double
    Dim remaining As Double
    Dim estimate As String
    Dim ws As Worksheet

    If gTotalItems = 0 Then Exit Sub
    percentComplete = currentIndex / gTotalItems
    elapsed = Timer - gStartTick

    If currentIndex > 0 Then
        remaining = (elapsed / currentIndex) * (gTotalItems - currentIndex)
    Else
        remaining = 0
    End If

    If remaining >= 60 Then
        estimate = Format(remaining / 60, "0") & " min"
    Else
        estimate = Format(remaining, "0") & " sec"
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    ws.Range(PROGRESS_PERCENT_CELL).Value = percentComplete
    ws.Range(PROGRESS_STATUS_CELL).Value = currentIndex & " / " & gTotalItems

    Application.StatusBar = Format(percentComplete, "0%") & " 완료 • 남은 시간 약 " & estimate & " • " & currentFile
End Sub

Public Sub FinishProgress(Optional ByVal success As Boolean = True)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    ws.Range(PROGRESS_LABEL_CELL).Value = IIf(success, "완료", "오류 발생")
    ws.Range(PROGRESS_STATUS_CELL).Value = ""
    Application.StatusBar = False
End Sub

Public Sub MarkFileStatus(ByVal rowIndex As Long, ByVal isSuccess As Boolean, ByVal message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    With ws.Cells(rowIndex, COL_STATUS)
        If isSuccess Then
            .Value = "Success"
            .Interior.Color = RGB(198, 239, 206)
        Else
            .Value = "Fail"
            .Interior.Color = RGB(255, 199, 206)
        End If
    End With
    ws.Cells(rowIndex, COL_MESSAGE).Value = message
End Sub

Public Sub ClearStatusColors()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    With ws.Range(ws.Cells(FILE_TABLE_START_ROW, COL_STATUS), ws.Cells(ws.Rows.Count, COL_STATUS))
        .Interior.ColorIndex = xlColorIndexNone
        .ClearContents
    End With
    ws.Range(ws.Cells(FILE_TABLE_START_ROW, COL_MESSAGE), ws.Cells(ws.Rows.Count, COL_MESSAGE)).ClearContents
End Sub

Public Function CreateLogFile() As String
    Dim fileName As String
    Dim folderPath As String

    folderPath = EnsureTrailingSlash(GetLogFolder)
    fileName = "SMB_" & Format(Now, "yyyymmdd_hhnnss") & ".log"
    CreateLogFile = folderPath & fileName

    With CreateObject("Scripting.FileSystemObject").CreateTextFile(CreateLogFile, True, True)
        .WriteLine "SMB Precheck Log"
        .WriteLine "Created: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        .WriteLine String(60, "-")
        .Close
    End With
End Function

Public Sub LogError(ByVal errorMessage As String, ByVal filePath As String)
    If gLogFilePath = vbNullString Then
        gLogFilePath = CreateLogFile()
    End If
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(gLogFilePath, 8, True, -1)
    stream.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & filePath & " | " & errorMessage
    stream.Close
End Sub

Public Sub ResetLogSession()
    gLogFilePath = vbNullString
End Sub

Public Function IsValidCellAddress(ByVal cellAddress As String) As Boolean
    Dim pattern As Object
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[A-Za-z]{1,3}[0-9]{1,7}$"
    pattern.IgnoreCase = True
    pattern.Global = False
    IsValidCellAddress = pattern.Test(Trim$(cellAddress))
End Function
