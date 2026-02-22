Attribute VB_Name = "modSorterShared"
Option Explicit

Public Const DEBUG_MODE As Boolean = False

Public Const SHEET_MAPPING As String = "mapping"
Public Const SHEET_TARGET As String = "target"
Public Const SHEET_MACRO As String = "log"
Public Const SHEET_TOOL As String = "macro"

Public Const TOBE_FILE_SUFFIX As String = "_TOBE"
Public Const START_ROW As Long = 5

' mapping sheet columns
Public Const COL_MAP_TOBE_TABLE As Long = 2   ' B
Public Const COL_MAP_TOBE_TABLE_LOGICAL As Long = 3 ' C
Public Const COL_MAP_TOBE_COLUMN As Long = 4  ' D
Public Const COL_MAP_TOBE_COLUMN_LOGICAL As Long = 5 ' E
Public Const COL_MAP_TOBE_KEY As Long = 6 ' F
Public Const COL_MAP_TOBE_DATA_TYPE As Long = 8 ' H
Public Const COL_MAP_TOBE_DIGITS As Long = 9 ' I
Public Const COL_MAP_TOBE_SCALE As Long = 10 ' J
Public Const COL_MAP_TOBE_DESC As Long = 12 ' L
Public Const COL_MAP_TOBE_NOTE As Long = 13 ' M
Public Const COL_MAP_ASIS_TABLE As Long = 15  ' O
Public Const COL_MAP_ASIS_COLUMN As Long = 17 ' Q

' target sheet columns
Public Const COL_TARGET_MAPPING_SHEET As Long = 2 ' B
Public Const COL_TARGET_FILE As Long = 3       ' C
Public Const COL_TARGET_SHEET As Long = 4      ' D
Public Const COL_TARGET_TOBE_TABLE As Long = 5 ' E
Public Const COL_TARGET_HEADER_CELL As Long = 6 ' F
Public Const COL_TARGET_DATA_CELL As Long = 7   ' G

' macro sheet input
Public Const MACRO_CFG_COL As Long = 2 ' B
Public Const MACRO_CFG_ROW_FILE_PATH As Long = 20
Public Const MACRO_CFG_ROW_TABLE_PHYS As Long = 21
Public Const MACRO_CFG_ROW_TABLE_LOGICAL As Long = 22
Public Const MACRO_CFG_ROW_COL_PHYS As Long = 23
Public Const MACRO_CFG_ROW_COL_LOGICAL As Long = 24
Public Const MACRO_CFG_ROW_KEY As Long = 25
Public Const MACRO_CFG_ROW_DATA_TYPE As Long = 26
Public Const MACRO_CFG_ROW_DIGITS As Long = 27
Public Const MACRO_CFG_ROW_SCALE As Long = 28
Public Const MACRO_CFG_ROW_DESC As Long = 29
Public Const MACRO_CFG_ROW_NOTE As Long = 30
Public Const MACRO_CFG_ROW_DELETE_FLAG As Long = 31
Public Const MACRO_CFG_ROW_SHEETS_START As Long = 32

' log area (B:D)
Public Const LOG_ROW_HEADER As Long = 1
Public Const LOG_ROW_START As Long = 2
Public Const LOG_COL_TIME As Long = 2
Public Const LOG_COL_LEVEL As Long = 3
Public Const LOG_COL_MESSAGE As Long = 4

Public Const MAX_EXCEL_ROW As Long = 1048576
Public Const MAX_EXCEL_COL As Long = 16384

Private gLogNextRow As Long
Private gPrevScreenUpdating As Boolean
Private gPrevEnableEvents As Boolean
Private gPrevDisplayAlerts As Boolean
Private gPrevCalculation As XlCalculation
Private gPrevDisplayStatusBar As Boolean
Private gPrevStatusBar As Variant
Private gLastStatusMessage As String
Private gLastDebugPoint As String
Public Function BuildMacroCellLabel(ByVal rowNum As Long) As String
    BuildMacroCellLabel = "macro!B" & CStr(rowNum)
End Function

Public Function JpRequiredSuffix() As String
    JpRequiredSuffix = " は必須です。"
End Function

Public Function JpAbsolutePathSuffix() As String
    JpAbsolutePathSuffix = " は絶対パスで入力してください: "
End Function

Public Function JpInvalidA1Suffix() As String
    JpInvalidA1Suffix = " のセル位置(A1形式)が不正です: "
End Function

Public Function JpNoSheetNamesMessage() As String
    JpNoSheetNamesMessage = "macro!B32以降にシート名を1件以上入力してください。"
End Function

Public Function JpDefSheetNotFoundSuffix() As String
    JpDefSheetNotFoundSuffix = " ToBeテーブル定義ファイルにシートが存在しません: "
End Function

Public Function JpDefFileNotFoundSuffix() As String
    JpDefFileNotFoundSuffix = " ToBeテーブル定義ファイルが存在しません: "
End Function

Public Function JpDuplicateSheetNameSuffix() As String
    JpDuplicateSheetNameSuffix = " シート名が重複しています: "
End Function

Public Function JpTablePhysEmptySuffix() As String
    JpTablePhysEmptySuffix = " テーブル物理名が空のため処理できません: "
End Function

Public Function JpSeedStartMessage() As String
    JpSeedStartMessage = "ToBeマッピング元ネタ作成支援を開始します。"
End Function

Public Function JpSeedCompletedMessage(ByVal outputSheetName As String, ByVal createdCount As Long) As String
    JpSeedCompletedMessage = "ToBeマッピング元ネタ作成支援が完了しました。" & vbCrLf & JpSeedOutputSheetPrefix() & outputSheetName & vbCrLf & JpSeedCreatedCountPrefix() & CStr(createdCount)
End Function

Public Function JpSeedOutputSheetPrefix() As String
    JpSeedOutputSheetPrefix = "出力シート名: "
End Function

Public Function JpSeedCreatedCountPrefix() As String
    JpSeedCreatedCountPrefix = "作成件数: "
End Function

Public Function JpSheetProcessedPrefix() As String
    JpSheetProcessedPrefix = "シート処理完了: "
End Function

Public Function JpSeedFatalPrefix() As String
    JpSeedFatalPrefix = "ToBeマッピング元ネタ作成支援でエラー"
End Function

Public Function JpSeedFatalDialogMessage() As String
    JpSeedFatalDialogMessage = "予期しないエラーが発生しました。logシートを確認してください。"
End Function

Public Function JpNoDataRowsSuffix() As String
    JpNoDataRowsSuffix = "にデータ行がありません。"
End Function

Public Function JpNoValidRowsSuffix() As String
    JpNoValidRowsSuffix = "に有効な行がありません。"
End Function

Public Function JpMappingAsIsPairError(ByVal mappingSheetName As String, ByVal rowNum As Long) As String
    JpMappingAsIsPairError = "[" & mappingSheetName & "] 行" & rowNum & "のAsIsテーブル(O列)とAsIsカラム(Q列)は両方入力または両方空欄にしてください。"
End Function

Public Function JpDuplicateToBeColumnError(ByVal mappingSheetName As String, ByVal rowNum As Long, ByVal toBeTable As String, ByVal toBeColumn As String) As String
    JpDuplicateToBeColumnError = "[" & mappingSheetName & "] 行" & rowNum & "でToBeカラムが重複しています: " & toBeTable & "." & toBeColumn
End Function

Public Function JpAsIsToToBeConflictError(ByVal mappingSheetName As String, ByVal rowNum As Long, ByVal asIsTable As String) As String
    JpAsIsToToBeConflictError = "[" & mappingSheetName & "] 行" & rowNum & "でAsIsテーブルの対応先が重複しています: " & asIsTable
End Function

Public Function JpToBeToAsIsConflictError(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    JpToBeToAsIsConflictError = "[" & mappingSheetName & "] ToBeテーブルに複数のAsIsテーブルが定義されています: " & toBeTable
End Function

Public Function JpToBeNoAsIsError(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    JpToBeNoAsIsError = "[" & mappingSheetName & "] ToBeテーブルにAsIsテーブルが1件も定義されていません: " & toBeTable
End Function

Public Function JpTargetDataStartRowOrderSuffix() As String
    JpTargetDataStartRowOrderSuffix = "ではデータ開始セル(G列)はカラム名開始セル(F列)より下の行を指定してください。"
End Function

Public Function JpTargetToBeTableMappingNotFoundError(ByVal rowNum As Long, ByVal toBeTable As String, ByVal mappingSheetName As String) As String
    JpTargetToBeTableMappingNotFoundError = TargetRowLabel(rowNum) & "でToBeテーブルに対応するマッピングが見つかりません: " & toBeTable & " (マッピングシート: " & mappingSheetName & ")"
End Function

Public Function JpTargetToBeMappingNotFoundError(ByVal rowNum As Long, ByVal mappingSheetName As String) As String
    JpTargetToBeMappingNotFoundError = TargetRowLabel(rowNum) & "でToBeテーブル定義が見つかりません (マッピングシート: " & mappingSheetName & ")。"
End Function

Public Function GetOrOpenWorkbook(ByVal filePath As String, ByVal readOnly As Boolean, ByRef openedBooks As Object, ByRef errors As Collection, ByVal targetRowNum As Long) As Workbook
    Dim key As String
    Dim wb As Workbook

    key = NormalizeKey(filePath)

    If openedBooks.Exists(key) Then
        DebugTrace "GetOrOpenWorkbook", "cache hit | row=" & CStr(targetRowNum) & ", file=" & filePath
        Set GetOrOpenWorkbook = openedBooks(key)
        Exit Function
    End If

    On Error GoTo OpenError
    DebugTrace "GetOrOpenWorkbook", "open | row=" & CStr(targetRowNum) & ", readOnly=" & CStr(readOnly) & ", file=" & filePath
    Set wb = Workbooks.Open(Filename:=filePath, readOnly:=readOnly, UpdateLinks:=0)
    openedBooks.Add key, wb
    Set GetOrOpenWorkbook = wb
    Exit Function

OpenError:
    DebugTrace "GetOrOpenWorkbook", "open error | row=" & CStr(targetRowNum) & ", file=" & filePath & ", err=" & CStr(Err.Number)
    AddError errors, "target 行" & targetRowNum & " のファイルを開けません: " & filePath & " / " & Err.Description
    Err.Clear
End Function

Public Function GetOrOpenWorkbookForRun(ByVal filePath As String, ByRef openedBooks As Object) As Workbook
    Dim key As String
    Dim wb As Workbook

    key = NormalizeKey(filePath)

    If openedBooks.Exists(key) Then
        DebugTrace "GetOrOpenWorkbookForRun", "cache hit | file=" & filePath
        Set GetOrOpenWorkbookForRun = openedBooks(key)
        Exit Function
    End If

    DebugTrace "GetOrOpenWorkbookForRun", "open | file=" & filePath
    Set wb = Workbooks.Open(Filename:=filePath, readOnly:=False, UpdateLinks:=0)
    openedBooks.Add key, wb
    Set GetOrOpenWorkbookForRun = wb
End Function

Public Function TryGetThisWorkbookSheet(ByVal sheetName As String, ByRef ws As Worksheet, ByRef errors As Collection) As Boolean
    Set ws = Nothing

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        AddError errors, "シート[" & sheetName & "]が見つかりません。"
        TryGetThisWorkbookSheet = False
    Else
        TryGetThisWorkbookSheet = True
    End If
End Function

Public Function TryGetWorksheetFromWorkbook(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    TryGetWorksheetFromWorkbook = Not (ws Is Nothing)
End Function

Public Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = TryGetWorksheetFromWorkbook(wb, sheetName, ws)
End Function

Public Sub CloseWorkbooks(ByVal openedBooks As Object, ByVal saveChanges As Boolean)
    Dim key As Variant

    On Error Resume Next
    For Each key In openedBooks.Keys
        openedBooks(key).Close saveChanges:=saveChanges
    Next key
    On Error GoTo 0
End Sub

Public Function NewDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set NewDictionary = dict
End Function

Public Sub AddError(ByRef errors As Collection, ByVal message As String)
    errors.Add message
End Sub

Public Sub ReportValidationErrors(ByRef errors As Collection)
    Dim item As Variant

    WriteLog LogLevelError(), "入力チェックでエラーが " & errors.Count & " 件見つかりました。"

    For Each item In errors
        WriteLog LogLevelError(), CStr(item)
    Next item

    MsgBox "入力チェックでエラーが " & errors.Count & " 件見つかりました。logシートを確認してください。", vbCritical + vbOKOnly
End Sub

Public Function TrimSafe(ByVal value As Variant) As String
    If IsError(value) Then
        TrimSafe = ""
    Else
        TrimSafe = Trim$(CStr(value))
    End If
End Function

Public Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = Trim$(value)
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    Dim found As String
    On Error Resume Next
    found = Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
    If Err.Number <> 0 Then
        DebugTrace "FileExists", "Dir error | file=" & filePath & ", err=" & CStr(Err.Number)
        Err.Clear
        FileExists = False
    Else
        FileExists = (Len(found) > 0)
    End If
    On Error GoTo 0
End Function

Public Function BuildToBeFilePath(ByVal sourceFilePath As String) As String
    Dim slashPosBack As Long
    Dim slashPosSlash As Long
    Dim slashPos As Long
    Dim folderPath As String
    Dim fileName As String
    Dim dotPos As Long
    Dim baseName As String
    Dim ext As String

    slashPosBack = InStrRev(sourceFilePath, "\")
    slashPosSlash = InStrRev(sourceFilePath, "/")

    If slashPosBack > slashPosSlash Then
        slashPos = slashPosBack
    Else
        slashPos = slashPosSlash
    End If

    If slashPos > 0 Then
        folderPath = Left(sourceFilePath, slashPos)
        fileName = Mid(sourceFilePath, slashPos + 1)
    Else
        folderPath = ""
        fileName = sourceFilePath
    End If

    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        baseName = Left(fileName, dotPos - 1)
        ext = Mid(fileName, dotPos)
    Else
        baseName = fileName
        ext = ""
    End If

    BuildToBeFilePath = folderPath & baseName & TOBE_FILE_SUFFIX & ext
End Function

Public Function BuildExecuteConfirmMessage() As String
    BuildExecuteConfirmMessage = "並び替えを実行しますか?"
End Function

Public Function LogLevelInfo() As String
    LogLevelInfo = "情報"
End Function

Public Function LogLevelError() As String
    LogLevelError = "エラー"
End Function

Public Sub DebugTrace(ByVal pointName As String, Optional ByVal detail As String = "")
    Dim message As String

    If Not DEBUG_MODE Then
        Exit Sub
    End If

    gLastDebugPoint = pointName
    If gLogNextRow < LOG_ROW_START Then
        Exit Sub
    End If

    message = "[DEBUG] " & pointName
    If Len(detail) > 0 Then
        message = message & " | " & detail
    End If
    WriteLog LogLevelInfo(), message
End Sub

Public Function LastDebugPointOrDefault() As String
    If Len(gLastDebugPoint) = 0 Then
        LastDebugPointOrDefault = "(なし)"
    Else
        LastDebugPointOrDefault = gLastDebugPoint
    End If
End Function

Public Function BuildUnexpectedErrorMessage(ByVal context As String, ByVal errNumber As Long, ByVal errDescription As String, ByVal errSource As String, Optional ByVal errLine As Long = 0) As String
    Dim sourceText As String
    Dim statusText As String
    Dim lineText As String
    Dim debugPointText As String

    sourceText = errSource
    If Len(sourceText) = 0 Then
        sourceText = "(なし)"
    End If

    statusText = gLastStatusMessage
    If Len(statusText) = 0 Then
        statusText = "(なし)"
    End If

    If errLine > 0 Then
        lineText = CStr(errLine)
    Else
        lineText = "(なし)"
    End If

    debugPointText = ""
    If DEBUG_MODE Then
        debugPointText = " [DebugPoint=" & LastDebugPointOrDefault() & "]"
    End If

    BuildUnexpectedErrorMessage = context & " 予期しないエラー: (" & CStr(errNumber) & ") " & errDescription & _
        " [Source=" & sourceText & "] [Line=" & lineText & "] [Status=" & statusText & "]" & debugPointText
End Function

Public Function TargetRowLabel(ByVal rowNum As Long) As String
    TargetRowLabel = "対象行" & CStr(rowNum)
End Function

Public Function ParseA1Address(ByVal address As String, ByRef colNumber As Long, ByRef rowNumber As Long) As Boolean
    Dim normalized As String
    Dim index As Long
    Dim currentChar As String
    Dim letters As String
    Dim digits As String

    ParseA1Address = False
    colNumber = 0
    rowNumber = 0

    normalized = UCase$(Trim$(address))
    If Len(normalized) = 0 Then
        Exit Function
    End If

    For index = 1 To Len(normalized)
        currentChar = Mid$(normalized, index, 1)

        If currentChar >= "A" And currentChar <= "Z" Then
            If Len(digits) > 0 Then
                Exit Function
            End If
            letters = letters & currentChar
        ElseIf currentChar >= "0" And currentChar <= "9" Then
            digits = digits & currentChar
        Else
            Exit Function
        End If
    Next index

    If Len(letters) = 0 Or Len(digits) = 0 Then
        Exit Function
    End If

    If Left$(digits, 1) = "0" Then
        Exit Function
    End If

    colNumber = ColumnLettersToNumber(letters)
    rowNumber = CLng(digits)

    If colNumber < 1 Or colNumber > MAX_EXCEL_COL Then
        Exit Function
    End If

    If rowNumber < 1 Or rowNumber > MAX_EXCEL_ROW Then
        Exit Function
    End If

    ParseA1Address = True
End Function

Public Function ColumnLettersToNumber(ByVal letters As String) As Long
    Dim index As Long
    Dim result As Long

    result = 0
    For index = 1 To Len(letters)
        result = result * 26 + (Asc(Mid$(letters, index, 1)) - Asc("A") + 1)
    Next index

    ColumnLettersToNumber = result
End Function

Public Function IsAbsolutePath(ByVal filePath As String) As Boolean
    Dim value As String

    value = Trim$(filePath)

    If value Like "[A-Za-z]:\*" Then
        IsAbsolutePath = True
        Exit Function
    End If

    If Left$(value, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If

    If Left$(value, 1) = "/" Then
        IsAbsolutePath = True
        Exit Function
    End If

    IsAbsolutePath = False
End Function

Public Function GetLastRowInColumns(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long) As Long
    Dim maxRow As Long
    maxRow = 1

    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col1).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col2).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col3).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col4).End(xlUp).Row)

    GetLastRowInColumns = maxRow
End Function

Public Function GetLastRowInColumns5(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long, ByVal col5 As Long) As Long
    Dim maxRow As Long
    maxRow = 1

    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col1).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col2).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col3).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col4).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col5).End(xlUp).Row)

    GetLastRowInColumns5 = maxRow
End Function

Public Function GetLastRowInColumns6(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long, ByVal col5 As Long, ByVal col6 As Long) As Long
    Dim maxRow As Long
    maxRow = 1

    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col1).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col2).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col3).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col4).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col5).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col6).End(xlUp).Row)

    GetLastRowInColumns6 = maxRow
End Function

Public Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range

    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = found.Row
    End If
End Function

Public Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim found As Range

    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = found.Column
    End If
End Function

Public Sub PrepareApplication()
    gPrevScreenUpdating = Application.ScreenUpdating
    gPrevEnableEvents = Application.EnableEvents
    gPrevDisplayAlerts = Application.DisplayAlerts
    gPrevCalculation = Application.Calculation
    gPrevDisplayStatusBar = Application.DisplayStatusBar
    gPrevStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    gLastStatusMessage = ""
    gLastDebugPoint = ""
End Sub

Public Sub RestoreApplication()
    On Error Resume Next
    Application.ScreenUpdating = gPrevScreenUpdating
    Application.EnableEvents = gPrevEnableEvents
    Application.DisplayAlerts = gPrevDisplayAlerts
    Application.Calculation = gPrevCalculation
    Application.StatusBar = gPrevStatusBar
    Application.DisplayStatusBar = gPrevDisplayStatusBar
    On Error GoTo 0
End Sub

Public Sub SetStatusMessage(ByVal message As String)
    On Error Resume Next
    gLastStatusMessage = message
    Application.DisplayStatusBar = True
    Application.StatusBar = message
    If DEBUG_MODE And gLogNextRow >= LOG_ROW_START Then
        WriteLog LogLevelInfo(), "[DEBUG] STATUS | " & message
    End If
    DoEvents
    On Error GoTo 0
End Sub

Public Sub SetStatusProgress(ByVal phase As String, ByVal current As Long, ByVal total As Long, Optional ByVal detail As String = "")
    Dim safeCurrent As Long
    Dim percentValue As Double
    Dim statusText As String

    If total <= 0 Then
        statusText = phase
    Else
        safeCurrent = current
        If safeCurrent < 0 Then safeCurrent = 0
        If safeCurrent > total Then safeCurrent = total
        percentValue = CDbl(safeCurrent) / CDbl(total)
        statusText = phase & " (" & CStr(safeCurrent) & "/" & CStr(total) & ", " & Format$(percentValue, "0%") & ")"
    End If

    If Len(detail) > 0 Then
        statusText = statusText & " - " & detail
    End If

    SetStatusMessage statusText
End Sub

Public Sub InitLog()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MACRO)
    On Error GoTo 0

    If ws Is Nothing Then
        gLogNextRow = LOG_ROW_START
        Exit Sub
    End If

    ws.Range(ws.Cells(LOG_ROW_START, LOG_COL_TIME), ws.Cells(ws.Rows.Count, LOG_COL_MESSAGE)).ClearContents

    ws.Cells(LOG_ROW_HEADER, LOG_COL_TIME).value = "日時"
    ws.Cells(LOG_ROW_HEADER, LOG_COL_LEVEL).value = "レベル"
    ws.Cells(LOG_ROW_HEADER, LOG_COL_MESSAGE).value = "メッセージ"

    gLogNextRow = LOG_ROW_START
End Sub

Public Sub WriteLog(ByVal level As String, ByVal message As String)
    Dim ws As Worksheet

    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets(SHEET_MACRO)

    ws.Cells(gLogNextRow, LOG_COL_TIME).value = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(gLogNextRow, LOG_COL_LEVEL).value = level
    ws.Cells(gLogNextRow, LOG_COL_MESSAGE).value = message
    gLogNextRow = gLogNextRow + 1
    Exit Sub

Fallback:
    Debug.Print Format$(Now, "yyyy/mm/dd hh:nn:ss") & " [" & level & "] " & message
End Sub
