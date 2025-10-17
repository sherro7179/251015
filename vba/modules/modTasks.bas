Attribute VB_Name = "modTasks"
Option Explicit

Public Function Task_UpdateFiles(ByVal wb As Workbook) As TaskResult
    On Error GoTo Fail_Handler

    Dim ws As Worksheet
    Dim rng As Range
    Dim mainName As String
    Dim depth1 As String
    Dim depth2 As String
    Dim i As Long
    Dim cell As Range

    Set ws = wb.Worksheets("Test Case")
    Set rng = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If rng.Rows.Count < 4 Then
        Err.Raise vbObjectError + 201, "Task_UpdateFiles", "Test Case 시트의 데이터 행이 충분하지 않습니다."
    End If

    mainName = Trim$(rng.Cells(1, 1).Value)
    If Len(mainName) = 0 Then
        Err.Raise vbObjectError + 202, "Task_UpdateFiles", "A2의 메인 식별자가 비어 있습니다."
    End If

    depth1 = mainName & "_00"
    depth2 = depth1 & "_01"
    rng.Cells(2, 1).Value = depth1
    rng.Cells(3, 1).Value = depth2

    For i = 5 To rng.Rows.Count
        Set cell = rng.Cells(i, 1)
        If Len(cell.Value) = 0 Then Exit For

        Select Case Len(cell.Value)
            Case Len(depth1)
                depth1 = IncrementLastNumber(depth1)
                cell.Value = depth1
                depth2 = depth1 & "_00"
            Case Len(depth2)
                If InStr(1, ws.Cells(i, "B").Value, "Precondition", vbTextCompare) > 0 Then
                    cell.Value = depth2
                Else
                    depth2 = IncrementLastNumber(depth2)
                    cell.Value = depth2
                End If
            Case Else
                Err.Raise vbObjectError + 203, "Task_UpdateFiles", "예상치 못한 ID 패턴입니다. 행: " & i
        End Select
    Next i

    Task_UpdateFiles.Success = True
    Task_UpdateFiles.Message = "케이스 ID 재정렬 완료"
    Exit Function

Fail_Handler:
    Task_UpdateFiles.Success = False
    Task_UpdateFiles.Message = Err.Description
End Function

Public Function Task_IOChange(ByVal wb As Workbook) As TaskResult
    On Error GoTo Fail_Handler

    Dim wsTestCase As Worksheet
    Dim wsIO As Worksheet
    Dim maxRow As Long
    Dim idx As Long
    Dim searchArea As Range

    Set wsTestCase = wb.Worksheets("Test Case")
    Set wsIO = ThisWorkbook.Worksheets(SHEET_IO_NAMES)
    maxRow = wsIO.Cells(wsIO.Rows.Count, "A").End(xlUp).Row
    If maxRow < 1 Then
        Err.Raise vbObjectError + 211, "Task_IOChange", "IO_name 시트에 치환 데이터가 없습니다."
    End If

    Set searchArea = wsTestCase.Range("A5:M700")
    For idx = 1 To maxRow
        If Len(wsIO.Cells(idx, 1).Value) > 0 Then
            searchArea.Replace What:=wsIO.Cells(idx, 1).Value, _
                               Replacement:=wsIO.Cells(idx, 2).Value, _
                               LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        End If
    Next idx

    Task_IOChange.Success = True
    Task_IOChange.Message = "IO 텍스트 치환 완료"
    Exit Function

Fail_Handler:
    Task_IOChange.Success = False
    Task_IOChange.Message = Err.Description
End Function

Public Function Task_ValueFind(ByVal wb As Workbook, ByVal findString As String, ByVal targetSheet As String, ByVal processedPath As String) As TaskResult
    On Error GoTo Fail_Handler

    Dim wsTestCase As Worksheet
    Dim wsUpdate As Worksheet
    Dim foundCell As Range
    Dim firstAddress As String
    Dim nextRow As Long
    Dim matchCount As Long

    Set wsTestCase = wb.Worksheets("Test Case")
    Set wsUpdate = ThisWorkbook.Worksheets(SHEET_DATA_UPDATE)

    If Len(findString) = 0 Then
        Err.Raise vbObjectError + 221, "Task_ValueFind", "찾을 문자열(B10)이 비어 있습니다."
    End If
    If Len(targetSheet) = 0 Then
        Err.Raise vbObjectError + 222, "Task_ValueFind", "대상 시트(B12)가 비어 있습니다."
    End If

    nextRow = wsUpdate.Cells(wsUpdate.Rows.Count, "A").End(xlUp).Row + 1
    Set foundCell = wsTestCase.Range("C:F").Find(What:=findString, LookAt:=xlPart, MatchCase:=False)

    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Do
            wsUpdate.Cells(nextRow, 1).Value = processedPath
            wsUpdate.Cells(nextRow, 2).Value = foundCell.Value
            wsUpdate.Cells(nextRow, 3).Value = targetSheet
            wsUpdate.Cells(nextRow, 4).Value = foundCell.Offset(0, 1).Value
            wsUpdate.Cells(nextRow, 5).Value = foundCell.Offset(0, 1).Address(False, False)
            nextRow = nextRow + 1
            matchCount = matchCount + 1
            Set foundCell = wsTestCase.Range("C:F").FindNext(foundCell)
        Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
    End If

    Task_ValueFind.Success = True
    If matchCount = 0 Then
        Task_ValueFind.Message = "일치 항목 없음"
    Else
        Task_ValueFind.Message = "일치 " & matchCount & "건"
    End If
    Exit Function

Fail_Handler:
    Task_ValueFind.Success = False
    Task_ValueFind.Message = Err.Description
End Function

Public Sub Task_ScriptMoveFolder(ByVal baseFolder As String)
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim fso As Object
    Dim subFolder As Object

    EnsureInfrastructure
    Set ws = ThisWorkbook.Worksheets(SHEET_SCRIPT_MOVE)
    ws.Range("A2:A" & ws.Rows.Count).ClearContents

    If Dir(baseFolder, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 231, "Task_ScriptMoveFolder", "폴더를 찾을 수 없습니다: " & baseFolder
    End If

    rowIndex = 2
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each subFolder In fso.GetFolder(baseFolder).SubFolders
        ws.Cells(rowIndex, 1).Value = EnsureTrailingSlash(subFolder.Path)
        rowIndex = rowIndex + 1
    Next subFolder
End Sub

Public Sub Task_ChangeValue()
    On Error GoTo Fail_Handler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim targetSheetName As String
    Dim cellAddress As String
    Dim newValue As Variant
    Dim wb As Workbook
    Dim targetSheet As Worksheet
    Dim successCount As Long
    Dim failCount As Long
    Dim displayName As String

    EnsureInfrastructure
    ResetLogSession

    Set ws = ThisWorkbook.Worksheets(SHEET_DATA_UPDATE)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "data_update 시트에 처리할 항목이 없습니다.", vbInformation
        Exit Sub
    End If

    StartProgress lastRow - 1, "값 일괄 변경"

    For i = 2 To lastRow
        filePath = ws.Cells(i, 1).Value
        targetSheetName = ws.Cells(i, 3).Value
        cellAddress = ws.Cells(i, 5).Value
        newValue = ws.Cells(i, 6).Value

        If Len(filePath) = 0 Or Len(targetSheetName) = 0 Or Len(cellAddress) = 0 Then GoTo NextIteration
        If Dir(filePath, vbNormal) = vbNullString Then
            failCount = failCount + 1
            LogError "파일을 찾을 수 없습니다.", filePath
            GoTo NextIteration
        End If

        Set wb = Workbooks.Open(filePath)
        On Error Resume Next
        Set targetSheet = wb.Worksheets(targetSheetName)
        On Error GoTo Fail_Handler
        If targetSheet Is Nothing Then
            failCount = failCount + 1
            LogError "대상 시트를 찾을 수 없습니다: " & targetSheetName, filePath
            wb.Close SaveChanges:=False
            Set targetSheet = Nothing
            GoTo NextIteration
        End If

        If Not IsValidCellAddress(cellAddress) Then
            failCount = failCount + 1
            LogError "유효하지 않은 셀 주소: " & cellAddress, filePath
            wb.Close SaveChanges:=False
            GoTo NextIteration
        End If

        targetSheet.Range(cellAddress).Value = newValue
        wb.Close SaveChanges:=True
        successCount = successCount + 1

NextIteration:
        If Len(filePath) > 0 Then
            displayName = Mid$(filePath, InStrRev(filePath, "\") + 1)
        Else
            displayName = "(no file)"
        End If
        UpdateProgress i - 1, displayName
    Next i

    FinishProgress True
    Dim message As String
    message = "Completed: " & successCount & " / " & (lastRow - 1)
    If failCount > 0 Then
        MsgBox message & vbCrLf & "실패: " & failCount & vbCrLf & "로그 파일: " & gLogFilePath, vbExclamation
    Else
        MsgBox message, vbInformation
    End If
    Exit Sub

Fail_Handler:
    FinishProgress False
    LogError Err.Description, "Task_ChangeValue"
    MsgBox "오류가 발생하여 작업이 중단되었습니다." & vbCrLf & Err.Description, vbCritical
End Sub

Public Function IncrementLastNumber(ByVal valueText As String) As String
    Dim pos As Long
    Dim num As Long

    pos = InStrRev(valueText, "_")
    If pos = 0 Then
        Err.Raise vbObjectError + 210, "IncrementLastNumber", "ID 형식이 올바르지 않습니다: " & valueText
    End If
    num = CLng(Mid$(valueText, pos + 1))
    num = num + 1
    IncrementLastNumber = Left$(valueText, pos) & Format$(num, "00")
End Function
