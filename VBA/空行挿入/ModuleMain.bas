Attribute VB_Name = "ModuleMain"
Option Explicit

Private Const VER = "1.0.0"
Private Const DEBUG_LOG_ENABLED As Boolean = False
Private Const LOG_FILE_NAME As String = "空行挿入_debug.log"

Private Type InstructionItem
    FilePath As String
    SheetName As String
    InsertCount As Long
    TargetRow As Long
    Direction As String
End Type

Public Sub 実行ボタン()
    On Error GoTo ErrHandler

    InitializeLog
    WriteLog "処理開始"

    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets("main")

    Dim currentRow As Long
    currentRow = 7

    Do While Trim$(wsMain.Cells(currentRow, "A").Value) <> vbNullString
        Dim instruction As InstructionItem
        instruction = ReadInstruction(wsMain, currentRow)

        WriteLog "指示処理開始: 行 " & currentRow & ", パス=" & instruction.FilePath
        ProcessInstruction instruction
        WriteLog "指示処理完了: 行 " & currentRow

        currentRow = currentRow + 1
    Loop

    WriteLog "処理正常終了"
    MsgBox "全ての処理が完了しました。", vbInformation
    Exit Sub

ErrHandler:
    WriteLog "致命的エラー: " & Err.Number & " - " & Err.Description
    MsgBox "エラーが発生しました。" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Function ReadInstruction(ByVal ws As Worksheet, ByVal rowIndex As Long) As InstructionItem
    On Error GoTo ErrHandler

    Dim item As InstructionItem

    item.FilePath = Trim$(ws.Cells(rowIndex, "A").Value)
    item.SheetName = Trim$(ws.Cells(rowIndex, "B").Value)
    item.InsertCount = ParsePositiveLong(ws.Cells(rowIndex, "C").Value, 1, "挿入する空行数", rowIndex)
    item.TargetRow = ParsePositiveLong(ws.Cells(rowIndex, "D").Value, 1, "挿入位置(行)", rowIndex)
    item.Direction = ParseDirection(ws.Cells(rowIndex, "E").Value, rowIndex)

    If item.FilePath = vbNullString Then
        Err.Raise vbObjectError + 100, , "Excelファイルパスが空です。行: " & rowIndex
    End If

    If Not IsSupportedExtension(item.FilePath) Then
        Err.Raise vbObjectError + 101, , "対応外の拡張子です。行: " & rowIndex & ", パス: " & item.FilePath
    End If

    If Dir$(item.FilePath) = vbNullString Then
        Err.Raise vbObjectError + 102, , "ファイルが存在しません。行: " & rowIndex & ", パス: " & item.FilePath
    End If

    ReadInstruction = item
    Exit Function

ErrHandler:
    WriteLog "指示読込エラー: 行 " & rowIndex & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ParsePositiveLong(ByVal value As Variant, ByVal defaultValue As Long, _
    ByVal columnLabel As String, ByVal rowIndex As Long) As Long

    Dim resultValue As Long

    If IsMissingValue(value) Then
        resultValue = defaultValue
    ElseIf IsNumeric(value) Then
        resultValue = CLng(value)
    Else
        Err.Raise vbObjectError + 103, , columnLabel & "が数値ではありません。行: " & rowIndex
    End If

    If resultValue < 1 Then
        Err.Raise vbObjectError + 104, , columnLabel & "は1以上である必要があります。行: " & rowIndex
    End If

    ParsePositiveLong = resultValue
End Function

Private Function ParseDirection(ByVal value As Variant, ByVal rowIndex As Long) As String
    Dim directionText As String
    directionText = Trim$(CStr(value))

    If directionText = vbNullString Then
        ParseDirection = "上"
        Exit Function
    End If

    Select Case directionText
        Case "上", "下"
            ParseDirection = directionText
        Case Else
            Err.Raise vbObjectError + 105, , "挿入方向は「上」または「下」を指定してください。行: " & rowIndex
    End Select
End Function

Private Function IsMissingValue(ByVal value As Variant) As Boolean
    IsMissingValue = _
        IsEmpty(value) Or _
        (VarType(value) = vbString And Trim$(CStr(value)) = vbNullString)
End Function

Private Sub ProcessInstruction(ByRef instruction As InstructionItem)
    On Error GoTo ErrHandler

    Dim targetWorkbook As Workbook
    Dim workbookOpenedHere As Boolean

    Set targetWorkbook = GetOpenWorkbookByFullName(instruction.FilePath)
    If targetWorkbook Is Nothing Then
        WriteLog "ブックを開きます: " & instruction.FilePath
        Set targetWorkbook = Workbooks.Open(instruction.FilePath, UpdateLinks:=False, ReadOnly:=False)
        workbookOpenedHere = True
    Else
        WriteLog "既存ブックを利用します: " & instruction.FilePath
    End If

    If instruction.SheetName = vbNullString Then
        Dim ws As Worksheet
        For Each ws In targetWorkbook.Worksheets
            ApplyInsertion ws, instruction
        Next ws
    Else
        If SheetExists(targetWorkbook, instruction.SheetName) Then
            ApplyInsertion targetWorkbook.Worksheets(instruction.SheetName), instruction
        Else
            Err.Raise vbObjectError + 106, , "シートが見つかりません。シート名: " & instruction.SheetName
        End If
    End If

    If workbookOpenedHere Then
        WriteLog "ブックを保存します: " & instruction.FilePath
        targetWorkbook.Save
        targetWorkbook.Close SaveChanges:=False
    Else
        WriteLog "ブックは既に開かれているため保存のみ実施: " & instruction.FilePath
        targetWorkbook.Save
    End If

    Exit Sub

ErrHandler:
    WriteLog "処理エラー: パス=" & instruction.FilePath & " - " & Err.Description
    If Not targetWorkbook Is Nothing Then
        If workbookOpenedHere Then
            On Error Resume Next
            targetWorkbook.Close SaveChanges:=False
            On Error GoTo 0
        End If
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ApplyInsertion(ByVal targetSheet As Worksheet, ByRef instruction As InstructionItem)
    WriteLog "シート処理: " & targetSheet.Name & ", 行=" & instruction.TargetRow & ", 行数=" & instruction.InsertCount & ", 方向=" & instruction.Direction

    Dim targetRow As Long
    targetRow = instruction.TargetRow

    Dim maxRow As Long
    maxRow = targetSheet.Rows.Count

    If targetRow < 1 Or targetRow > maxRow Then
        Err.Raise vbObjectError + 107, , "挿入位置(行)がシートの範囲外です。シート: " & targetSheet.Name & ", 行: " & targetRow
    End If

    If instruction.Direction = "下" And targetRow = maxRow Then
        Err.Raise vbObjectError + 108, , "挿入方向「下」は最終行では使用できません。シート: " & targetSheet.Name & ", 行: " & targetRow
    End If

    If instruction.Direction = "上" Then
        targetSheet.Rows(targetRow).Resize(instruction.InsertCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        targetSheet.Rows(targetRow + 1).Resize(instruction.InsertCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
End Sub

Private Function SheetExists(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In targetWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function

Private Function GetOpenWorkbookByFullName(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function IsSupportedExtension(ByVal filePath As String) As Boolean
    Dim extensionValue As String
    extensionValue = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
    IsSupportedExtension = (extensionValue = "xlsx" Or extensionValue = "xlsm")
End Function

Private Sub InitializeLog()
    If Not DEBUG_LOG_ENABLED Then
        Exit Sub
    End If

    Dim logPath As String
    logPath = GetLogFilePath

    On Error Resume Next
    Kill logPath
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal message As String)
    If Not DEBUG_LOG_ENABLED Then
        Exit Sub
    End If

    Dim logPath As String
    logPath = GetLogFilePath

    Dim fileNum As Integer
    fileNum = FreeFile

    On Error GoTo ErrHandler
    Open logPath For Append As #fileNum
    Print #fileNum, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " - " & message
    Close #fileNum
    Exit Sub

ErrHandler:
    Close #fileNum
End Sub

Private Function GetLogFilePath() As String
    GetLogFilePath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE_NAME
End Function
