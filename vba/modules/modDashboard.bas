Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub Command_SelectFolderPath()
    On Error GoTo Fail_Handler

    Dim dialog As FileDialog
    Dim folderPath As String
    Dim ws As Worksheet
    Dim fileCount As Long

    EnsureInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)

    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "테스트 케이스 폴더 선택"

    If dialog.Show = -1 Then
        folderPath = EnsureTrailingSlash(dialog.SelectedItems(1))
        ws.Range(BASE_PATH_CELL).Value = folderPath
        fileCount = ListExcelFilesInFolder(folderPath)
        MsgBox fileCount & "개의 파일을 목록에 로드했습니다.", vbInformation
    End If
    Exit Sub

Fail_Handler:
    FinishProgress False
    MsgBox "폴더를 불러오는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub Command_UpdateFiles()
    RunSelectedFilesTask taskUpdateFiles, "케이스 ID 재번호"
End Sub

Public Sub Command_IOChange()
    RunSelectedFilesTask taskIOChange, "IO 텍스트 치환"
End Sub

Public Sub Command_ValueFind()
    Dim ws As Worksheet
    EnsureInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_DATA_UPDATE)
    ws.Range("A2:F" & ws.Rows.Count).ClearContents
    RunSelectedFilesTask taskValueFind, "값 찾기"
End Sub

Public Sub Command_ChangeValue()
    Task_ChangeValue
End Sub

Public Sub Command_ScriptMoveFolderPath()
    On Error GoTo Fail_Handler
    Dim ws As Worksheet
    Dim basePath As String

    EnsureInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_FILES)
    basePath = EnsureTrailingSlash(ws.Range(BASE_PATH_CELL).Value)
    If Len(basePath) = 0 Then
        Err.Raise vbObjectError + 301, "Command_ScriptMoveFolderPath", "B2 셀에 경로가 지정되지 않았습니다."
    End If
    If Dir(basePath, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 302, "Command_ScriptMoveFolderPath", "지정한 폴더를 찾을 수 없습니다."
    End If

    Task_ScriptMoveFolder basePath
    MsgBox "하위 폴더 목록을 script_move 시트에 업데이트했습니다.", vbInformation
    Exit Sub

Fail_Handler:
    MsgBox "하위 폴더를 가져오는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub RunSelectedFilesTask(ByVal operation As TaskOperation, ByVal taskLabel As String)
    On Error GoTo Fail_Handler

    Dim wsFiles As Worksheet
    Dim basePath As String
    Dim entries() As SelectedFileEntry
    Dim countSelected As Long
    Dim idx As Long
    Dim processedPath As String
    Dim wb As Workbook
    Dim result As TaskResult
    Dim successList As String
    Dim successCount As Long
    Dim findString As String
    Dim targetSheet As String
    Dim saveChanges As Boolean

    EnsureInfrastructure
    ResetLogSession
    ClearStatusColors

    Set wsFiles = ThisWorkbook.Worksheets(SHEET_FILES)
    basePath = EnsureTrailingSlash(wsFiles.Range(BASE_PATH_CELL).Value)
    If Len(basePath) = 0 Then
        Err.Raise vbObjectError + 311, "RunSelectedFilesTask", "폴더 경로가 지정되지 않았습니다."
    End If
    If Dir(basePath, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 312, "RunSelectedFilesTask", "폴더를 찾을 수 없습니다: " & basePath
    End If

    countSelected = GetSelectedFiles(basePath, entries)
    If countSelected = 0 Then
        wsFiles.Parent.Save
        MsgBox "선택된 파일이 없습니다. Include? 열 또는 필터를 확인하세요.", vbExclamation
        Exit Sub
    End If

    If operation = taskValueFind Then
        findString = wsFiles.Range(FIND_VALUE_CELL).Value
        targetSheet = wsFiles.Range(TARGET_SHEET_CELL).Value
    End If

    StartProgress countSelected, taskLabel

    For idx = 1 To countSelected
        UpdateProgress idx - 1, entries(idx).FileName
        processedPath = PrepareProcessedFile(entries(idx).OriginalPath)
        entries(idx).ProcessedPath = processedPath

        Set wb = Workbooks.Open(processedPath, UpdateLinks:=False, ReadOnly:=False)
        saveChanges = False
        Select Case operation
            Case taskUpdateFiles
                result = Task_UpdateFiles(wb)
                saveChanges = True
            Case taskIOChange
                result = Task_IOChange(wb)
                saveChanges = True
            Case taskValueFind
                result = Task_ValueFind(wb, findString, targetSheet, processedPath)
        End Select

        If result.Success Then
            If saveChanges Then
                wb.Close SaveChanges:=True
            Else
                wb.Close SaveChanges:=False
            End If
            successList = successList & entries(idx).FileName & vbCrLf
            successCount = successCount + 1
            MarkFileStatus entries(idx).RowIndex, True, result.Message
            UpdateProgress idx, entries(idx).FileName
        Else
            wb.Close SaveChanges:=False
            LogError result.Message, entries(idx).OriginalPath
            MarkFileStatus entries(idx).RowIndex, False, result.Message
            FinishProgress False
            MsgBox "오류가 발생하여 작업이 중단되었습니다." & vbCrLf & _
                   "파일: " & entries(idx).FileName & vbCrLf & result.Message, vbCritical
            Exit Sub
        End If
    Next idx

    FinishProgress True

    Dim summary As String
    summary = taskLabel & " 완료" & vbCrLf & _
              "성공: " & successCount & " / " & countSelected
    If gLogFilePath <> vbNullString Then
        summary = summary & vbCrLf & "로그: " & gLogFilePath
    End If
    summary = summary & vbCrLf & vbCrLf & "성공 파일 목록:" & vbCrLf & successList
    MsgBox summary, vbInformation
    Exit Sub

Fail_Handler:
    FinishProgress False
    LogError Err.Description, "RunSelectedFilesTask"
    MsgBox "작업을 시작하지 못했습니다." & vbCrLf & Err.Description, vbCritical
End Sub
