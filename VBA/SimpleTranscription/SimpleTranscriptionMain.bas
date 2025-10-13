Attribute VB_Name = "SimpleTranscriptionMain"
Option Explicit

Private Enum NotFoundBehavior
    NotFoundIgnore = 0
    NotFoundAbort = 1
End Enum

Private Type AppConfig
    NotFoundMode As NotFoundBehavior
    KeepWorkbooksOpen As Boolean
    SkipEmptySource As Boolean
    CaseSensitive As Boolean
    AllowPartialMatch As Boolean
    DistinguishWidth As Boolean
    WriteAsComment As Boolean
End Type

Private Type TransferInstruction
    RowIndex As Long
    SourcePath As String
    SourceSheetName As String
    SourceSearchColumn As Long
    SourceTransferColumn As Long
    DestPath As String
    DestSheetName As String
    DestSearchColumn As Long
    DestTransferColumn As Long
End Type

Private Const MAIN_SHEET_NAME As String = "main"
Private Const FIRST_TRANSFER_ROW As Long = 19

Private Const COL_STATUS As String = "D"
Private Const COL_SOURCE_FILE As String = "E"
Private Const COL_SOURCE_SHEET As String = "F"
Private Const COL_SOURCE_SEARCH As String = "G"
Private Const COL_SOURCE_TRANSFER As String = "H"
Private Const COL_DEST_FILE As String = "I"
Private Const COL_DEST_SHEET As String = "J"
Private Const COL_DEST_SEARCH As String = "K"
Private Const COL_DEST_TRANSFER As String = "L"

Private Const SETTING_NOT_FOUND As String = "H6"
Private Const SETTING_KEEP_OPEN As String = "H7"
Private Const SETTING_SKIP_EMPTY As String = "H8"
Private Const SETTING_CASE_SENSITIVE As String = "H9"
Private Const SETTING_ALLOW_PARTIAL As String = "H10"
Private Const SETTING_DISTINGUISH_WIDTH As String = "H11"
Private Const SETTING_AS_COMMENT As String = "H12"

Public Sub Run_Click()
    Const PROC_NAME As String = "Run_Click"
    
    On Error GoTo ErrHandler
    
    Dim mainSheet As Worksheet
    Set mainSheet = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    
    Dim logPath As String
    logPath = BuildLogPath()
    SimpleTranscriptionLogging.StartLog logPath
    SimpleTranscriptionLogging.WriteLog PROC_NAME & " : 処理開始"
    
    Dim config As AppConfig
    config = LoadAppConfig(mainSheet)
    Dim configLoaded As Boolean
    configLoaded = True
    
    Dim keepWorkbooksOpen As Boolean
    keepWorkbooksOpen = config.KeepWorkbooksOpen
    
    Dim openedByMacro As Object
    Set openedByMacro = CreateObject("Scripting.Dictionary")
    
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousStateCaptured As Boolean
    Dim applicationStateChanged As Boolean
    
    previousStateCaptured = False
    applicationStateChanged = False
    
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts
    previousStateCaptured = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    applicationStateChanged = True
    
    ProcessTransfers mainSheet, config, openedByMacro
    
    SimpleTranscriptionLogging.WriteLog PROC_NAME & " : 処理正常終了"
    MsgBox "転記が完了しました。", vbInformation + vbOKOnly, "シンプル転記"
    
Cleanup:
    On Error Resume Next
    If previousStateCaptured And applicationStateChanged Then
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEnableEvents
        Application.DisplayAlerts = previousDisplayAlerts
    End If
    
    If Not openedByMacro Is Nothing Then
        If configLoaded And keepWorkbooksOpen = False Then
            CloseWorkbooks openedByMacro
        End If
    End If
    
    SimpleTranscriptionLogging.StopLog
    Exit Sub
    
ErrHandler:
    SimpleTranscriptionLogging.WriteLog PROC_NAME & " : エラー発生 " & Err.Number & " - " & Err.Description
    MsgBox "エラーが発生しました。" & vbCrLf & Err.Description, vbCritical + vbOKOnly, "シンプル転記"
    Resume Cleanup
End Sub

Private Sub ProcessTransfers(ByVal mainSheet As Worksheet, ByRef config As AppConfig, ByVal openedByMacro As Object)
    Dim currentRow As Long
    currentRow = FIRST_TRANSFER_ROW
    
    Do
        Dim statusValue As String
        statusValue = Trim$(UCase$(CStr(mainSheet.Range(COL_STATUS & currentRow).Value)))
        
        If Len(statusValue) = 0 Then
            SimpleTranscriptionLogging.WriteLog "行" & currentRow & " : 空セル検知のため処理終了"
            Exit Do
        End If
        
        If statusValue = "STOPPER" Then
            SimpleTranscriptionLogging.WriteLog "行" & currentRow & " : STOPPER検知のため処理終了"
            Exit Do
        End If
        
        If statusValue = "DISABLE" Then
            SimpleTranscriptionLogging.WriteLog "行" & currentRow & " : DISABLEのためスキップ"
        ElseIf statusValue = "ENABLE" Then
            Dim instruction As TransferInstruction
            instruction = LoadInstruction(mainSheet, currentRow)
            SimpleTranscriptionLogging.WriteLog "行" & currentRow & " : 転記処理開始"
            ExecuteInstruction instruction, config, openedByMacro
            SimpleTranscriptionLogging.WriteLog "行" & currentRow & " : 転記処理終了"
        Else
            Err.Raise vbObjectError + 100, , "main!" & COL_STATUS & currentRow & " の値が不正です: " & statusValue
        End If
        
        currentRow = currentRow + 1
    Loop
End Sub

Private Function LoadAppConfig(ByVal mainSheet As Worksheet) As AppConfig
    Dim config As AppConfig
    
    config.NotFoundMode = ParseNotFoundBehavior(mainSheet.Range(SETTING_NOT_FOUND).Value)
    config.KeepWorkbooksOpen = ParseYesNo(mainSheet.Range(SETTING_KEEP_OPEN).Value, True)
    config.SkipEmptySource = ParseYesNo(mainSheet.Range(SETTING_SKIP_EMPTY).Value, True)
    config.CaseSensitive = ParseYesNo(mainSheet.Range(SETTING_CASE_SENSITIVE).Value, False)
    config.AllowPartialMatch = ParseYesNo(mainSheet.Range(SETTING_ALLOW_PARTIAL).Value, False)
    config.DistinguishWidth = ParseYesNo(mainSheet.Range(SETTING_DISTINGUISH_WIDTH).Value, False)
    config.WriteAsComment = ParseYesNo(mainSheet.Range(SETTING_AS_COMMENT).Value, False)
    
    LoadAppConfig = config
End Function

Private Function ParseNotFoundBehavior(ByVal settingValue As Variant) As NotFoundBehavior
    Dim textValue As String
    textValue = Trim$(CStr(settingValue))
    
    If Len(textValue) = 0 Then
        ParseNotFoundBehavior = NotFoundIgnore
        Exit Function
    End If
    
    Select Case UCase$(textValue)
        Case "無視", "IGNORE"
            ParseNotFoundBehavior = NotFoundIgnore
        Case "中断", "STOP"
            ParseNotFoundBehavior = NotFoundAbort
        Case Else
            Err.Raise vbObjectError + 101, , "設定表(H6)の値が不正です: " & textValue
    End Select
End Function

Private Function ParseYesNo(ByVal settingValue As Variant, ByVal defaultValue As Boolean) As Boolean
    If IsEmpty(settingValue) Then
        ParseYesNo = defaultValue
        Exit Function
    End If
    
    Dim textValue As String
    textValue = Trim$(UCase$(CStr(settingValue)))
    
    Select Case textValue
        Case ""
            ParseYesNo = defaultValue
        Case "YES"
            ParseYesNo = True
        Case "NO"
            ParseYesNo = False
        Case Else
            Err.Raise vbObjectError + 102, , "設定表のYES/NOに不正な値があります: " & textValue
    End Select
End Function

Private Function LoadInstruction(ByVal mainSheet As Worksheet, ByVal rowIndex As Long) As TransferInstruction
    Dim info As TransferInstruction
    
    info.RowIndex = rowIndex
    info.SourcePath = Trim$(CStr(mainSheet.Range(COL_SOURCE_FILE & rowIndex).Value))
    info.SourceSheetName = Trim$(CStr(mainSheet.Range(COL_SOURCE_SHEET & rowIndex).Value))
    info.SourceSearchColumn = ColumnLetterToNumber(mainSheet.Range(COL_SOURCE_SEARCH & rowIndex).Value)
    info.SourceTransferColumn = ColumnLetterToNumber(mainSheet.Range(COL_SOURCE_TRANSFER & rowIndex).Value)
    info.DestPath = Trim$(CStr(mainSheet.Range(COL_DEST_FILE & rowIndex).Value))
    info.DestSheetName = Trim$(CStr(mainSheet.Range(COL_DEST_SHEET & rowIndex).Value))
    info.DestSearchColumn = ColumnLetterToNumber(mainSheet.Range(COL_DEST_SEARCH & rowIndex).Value)
    info.DestTransferColumn = ColumnLetterToNumber(mainSheet.Range(COL_DEST_TRANSFER & rowIndex).Value)
    
    ValidateInstruction info
    
    LoadInstruction = info
End Function

Private Sub ValidateInstruction(ByRef info As TransferInstruction)
    If Len(info.SourcePath) = 0 Then
        Err.Raise vbObjectError + 110, , "行" & info.RowIndex & " : 転記元ファイル名が未入力です。"
    End If
    
    If Len(info.SourceSheetName) = 0 Then
        Err.Raise vbObjectError + 111, , "行" & info.RowIndex & " : 転記元シート名が未入力です。"
    End If
    
    If info.SourceSearchColumn = 0 Then
        Err.Raise vbObjectError + 112, , "行" & info.RowIndex & " : 転記元検索列が不正です。"
    End If
    
    If info.SourceTransferColumn = 0 Then
        Err.Raise vbObjectError + 113, , "行" & info.RowIndex & " : 転記元転記列が不正です。"
    End If
    
    If Len(info.DestPath) = 0 Then
        Err.Raise vbObjectError + 114, , "行" & info.RowIndex & " : 転記先ファイル名が未入力です。"
    End If
    
    If Len(info.DestSheetName) = 0 Then
        Err.Raise vbObjectError + 115, , "行" & info.RowIndex & " : 転記先シート名が未入力です。"
    End If
    
    If info.DestSearchColumn = 0 Then
        Err.Raise vbObjectError + 116, , "行" & info.RowIndex & " : 転記先検索列が不正です。"
    End If
    
    If info.DestTransferColumn = 0 Then
        Err.Raise vbObjectError + 117, , "行" & info.RowIndex & " : 転記先転記列が不正です。"
    End If
End Sub

Private Sub ExecuteInstruction(ByRef instruction As TransferInstruction, ByRef config As AppConfig, ByVal openedByMacro As Object)
    Dim sourceBook As Workbook
    Set sourceBook = GetOrOpenWorkbook(instruction.SourcePath, openedByMacro)
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = GetWorksheetByName(sourceBook, instruction.SourceSheetName, "転記元", instruction.RowIndex)
    
    Dim lastSourceRow As Long
    lastSourceRow = GetLastUsedRow(sourceSheet, instruction.SourceSearchColumn)
    
    Dim yellowFound As Long
    yellowFound = 0
    
    Dim rowPointer As Long
    For rowPointer = 1 To lastSourceRow
        Dim searchCell As Range
        Set searchCell = sourceSheet.Cells(rowPointer, instruction.SourceSearchColumn)
        
        If IsCellYellow(searchCell) Then
            yellowFound = yellowFound + 1
            
            Dim searchValue As Variant
            searchValue = searchCell.Value
            
            If IsBlankValue(searchValue) Then
                SimpleTranscriptionLogging.WriteLog "行" & instruction.RowIndex & " : 転記元行" & rowPointer & " の検索値が空のためスキップ"
                GoTo ContinueLoop
            End If
            
            Dim transferValue As Variant
            transferValue = sourceSheet.Cells(rowPointer, instruction.SourceTransferColumn).Value
            
            If config.SkipEmptySource And IsBlankValue(transferValue) Then
                SimpleTranscriptionLogging.WriteLog "行" & instruction.RowIndex & " : 転記元行" & rowPointer & " の転記値が空のためスキップ"
                GoTo ContinueLoop
            End If
            
            ApplyTransfer instruction, config, openedByMacro, searchValue, transferValue, rowPointer
        End If
ContinueLoop:
    Next rowPointer
    
    If yellowFound = 0 Then
        SimpleTranscriptionLogging.WriteLog "行" & instruction.RowIndex & " : 転記元シートに黄色セルがありません"
    End If
End Sub

Private Sub ApplyTransfer(ByRef instruction As TransferInstruction, ByRef config As AppConfig, ByVal openedByMacro As Object, ByVal searchValue As Variant, ByVal transferValue As Variant, ByVal sourceRowIndex As Long)
    Dim destBook As Workbook
    Set destBook = GetOrOpenWorkbook(instruction.DestPath, openedByMacro)
    
    Dim destSheet As Worksheet
    Set destSheet = GetWorksheetByName(destBook, instruction.DestSheetName, "転記先", instruction.RowIndex)
    
    Dim targetRow As Long
    targetRow = FindDestinationRow(destSheet, instruction.DestSearchColumn, searchValue, config)
    
    If targetRow = 0 Then
        Dim message As String
        message = "行" & instruction.RowIndex & " : " & _
                  "転記先シート """ & destSheet.Name & """ の検索列で一致が見つかりません。" & _
                  " 検索値=" & CStr(searchValue)
        SimpleTranscriptionLogging.WriteLog message
        If config.NotFoundMode = NotFoundAbort Then
            Err.Raise vbObjectError + 120, , message
        End If
        Exit Sub
    End If
    
    Dim destinationCell As Range
    Set destinationCell = destSheet.Cells(targetRow, instruction.DestTransferColumn)
    
    If config.WriteAsComment Then
        ApplyAsComment destinationCell, transferValue
    Else
        destinationCell.Value = transferValue
    End If
    
    SimpleTranscriptionLogging.WriteLog "行" & instruction.RowIndex & " : 転記先行" & targetRow & " へ書き込みました"
End Sub

Private Function GetWorksheetByName(ByVal targetBook As Workbook, ByVal sheetName As String, ByVal roleName As String, ByVal rowIndex As Long) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = targetBook.Worksheets(sheetName)
    On Error GoTo 0
    
    If GetWorksheetByName Is Nothing Then
        Err.Raise vbObjectError + 130, , "行" & rowIndex & " : " & roleName & "ファイル """ & targetBook.FullName & """ にシート """ & sheetName & """ が見つかりません。"
    End If
End Function

Private Function GetLastUsedRow(ByVal targetSheet As Worksheet, ByVal columnIndex As Long) As Long
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, columnIndex).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    GetLastUsedRow = lastRow
End Function

Private Function FindDestinationRow(ByVal destSheet As Worksheet, ByVal columnIndex As Long, ByVal searchValue As Variant, ByRef config As AppConfig) As Long
    Dim lastRow As Long
    lastRow = GetLastUsedRow(destSheet, columnIndex)
    
    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        Dim candidate As Variant
        candidate = destSheet.Cells(rowIndex, columnIndex).Value
        If StringsMatch(searchValue, candidate, config.AllowPartialMatch, config.CaseSensitive, config.DistinguishWidth) Then
            FindDestinationRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub ApplyAsComment(ByVal destinationCell As Range, ByVal transferValue As Variant)
    On Error Resume Next
    destinationCell.ClearComments
    destinationCell.Comment.Delete
    On Error GoTo 0
    
    Dim textValue As String
    If IsError(transferValue) Then
        textValue = ""
    Else
        textValue = CStr(transferValue)
    End If
    
    On Error Resume Next
    destinationCell.AddComment textValue
    If Err.Number <> 0 Then
        Err.Clear
        destinationCell.AddCommentThreaded textValue
    End If
    On Error GoTo 0
End Sub

Private Function GetOrOpenWorkbook(ByVal fileNameOrPath As String, ByVal openedByMacro As Object) As Workbook
    Dim fullPath As String
    fullPath = ResolveWorkbookPath(fileNameOrPath)
    
    If Len(fullPath) = 0 Then
        Err.Raise vbObjectError + 140, , "ファイルが見つかりません: " & fileNameOrPath
    End If
    
    Dim key As String
    key = LCase$(fullPath)
    
    Dim existingWorkbook As Workbook
    Set existingWorkbook = FindOpenWorkbook(fullPath)
    
    If Not existingWorkbook Is Nothing Then
        If Not openedByMacro.Exists(key) Then
            openedByMacro.Add key, False
        End If
        Set GetOrOpenWorkbook = existingWorkbook
        Exit Function
    End If
    
    If FileExists(fullPath) = False Then
        Err.Raise vbObjectError + 141, , "ファイルが存在しません: " & fullPath
    End If
    
    SimpleTranscriptionLogging.WriteLog "ファイルを開きます: " & fullPath
    
    Set GetOrOpenWorkbook = Application.Workbooks.Open(fullPath, UpdateLinks:=False, ReadOnly:=False)
    
    If openedByMacro.Exists(key) Then
        openedByMacro(key) = True
    Else
        openedByMacro.Add key, True
    End If
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit For
        End If
    Next wb
End Function

Private Function FileExists(ByVal pathName As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(pathName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "")
    On Error GoTo 0
End Function

Private Sub CloseWorkbooks(ByVal openedByMacro As Object)
    Dim key As Variant
    For Each key In openedByMacro.Keys
        If openedByMacro(key) = True Then
            Dim fullPath As String
            fullPath = key
            
            If StrComp(ThisWorkbook.FullName, fullPath, vbTextCompare) = 0 Then
                GoTo ContinueLoop
            End If
            
            Dim wb As Workbook
            Set wb = FindOpenWorkbook(fullPath)
            If Not wb Is Nothing Then
                SimpleTranscriptionLogging.WriteLog "ファイルを閉じます: " & fullPath
                On Error Resume Next
                wb.Close SaveChanges:=True
                On Error GoTo 0
            End If
        End If
ContinueLoop:
    Next key
End Sub

Private Function BuildLogPath() As String
    Dim basePath As String
    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then Exit Function
    BuildLogPath = basePath & Application.PathSeparator & "SimpleTranscription_debug.log"
End Function
