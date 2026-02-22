Attribute VB_Name = "modSorterSort"
Option Explicit

Private Const KEY_TARGET_ROWNUM As String = "RowNum"
Private Const KEY_TARGET_MAPPING_SHEET As String = "MappingSheetName"
Private Const KEY_TARGET_FILE_PATH As String = "FilePath"
Private Const KEY_TARGET_SHEET_NAME As String = "SheetName"
Private Const KEY_TARGET_TOBE_TABLE As String = "ToBeTable"
Private Const KEY_TARGET_HEADER_CELL As String = "HeaderCell"
Private Const KEY_TARGET_DATA_CELL As String = "DataCell"
Private Const KEY_TARGET_OUTPUT_FILE_PATH As String = "OutputFilePath"
Private Const KEY_TARGET_TOBE_TABLE_KEY As String = "ToBeTableKey"
Private Const KEY_PLAN_SOURCE_FILE As String = "SourceFilePath"
Private Const KEY_PLAN_OUTPUT_FILE As String = "OutputFilePath"
Private Const KEY_PLAN_TARGET_SHEETS As String = "TargetSheets"

Private Const KEY_MAP_ROWNUM As String = "RowNum"
Private Const KEY_MAP_TOBE_TABLE As String = "ToBeTable"
Private Const KEY_MAP_TOBE_COLUMN As String = "ToBeColumn"
Private Const KEY_MAP_ASIS_TABLE As String = "AsIsTable"
Private Const KEY_MAP_ASIS_COLUMN As String = "AsIsColumn"
Private Const PROGRESS_INTERVAL_ROWS As Long = 50


Private Type TargetLineData
    MappingSheetName As String
    FilePath As String
    SheetName As String
    ToBeTable As String
    HeaderCell As String
    DataCell As String
End Type
Public Sub StartSort()
    Dim errors As Collection
    Dim mappingByToBe As Object
    Dim asIsToToBe As Object
    Dim targets As Collection

    On Error GoTo FatalError

    If MsgBox(BuildExecuteConfirmMessage(), vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
        Exit Sub
    End If

    PrepareApplication
    InitLog
    WriteLog LogLevelInfo(), "並び替え処理を開始します。"
    DebugTrace "StartSort", "開始"

    Set errors = New Collection
    Set mappingByToBe = NewDictionary()
    Set asIsToToBe = NewDictionary()
    Set targets = New Collection

    SetStatusMessage "並び替え: target入力をチェック中"
    DebugTrace "StartSort", "ValidateTarget 呼出"
    ValidateTarget targets, errors
    DebugTrace "StartSort", "ValidateTarget 完了 | targets=" & CStr(targets.Count) & ", errors=" & CStr(errors.Count)
    SetStatusMessage "並び替え: mapping定義をチェック中"
    DebugTrace "StartSort", "ValidateMappingsForTargets 呼出"
    ValidateMappingsForTargets targets, mappingByToBe, asIsToToBe, errors
    DebugTrace "StartSort", "ValidateMappingsForTargets 完了 | mapCount=" & CStr(mappingByToBe.Count) & ", errors=" & CStr(errors.Count)
    SetStatusMessage "並び替え: targetとmappingを紐づけ中"
    DebugTrace "StartSort", "BindTargetsToMapping 呼出"
    BindTargetsToMapping targets, mappingByToBe, errors
    DebugTrace "StartSort", "BindTargetsToMapping 完了 | errors=" & CStr(errors.Count)
    SetStatusMessage "並び替え: ファイル・シート・カラムを事前チェック中"
    DebugTrace "StartSort", "ValidateFilesSheetsAndColumns 呼出"
    ValidateFilesSheetsAndColumns targets, mappingByToBe, errors
    DebugTrace "StartSort", "ValidateFilesSheetsAndColumns 完了 | errors=" & CStr(errors.Count)

    If errors.Count > 0 Then
        ReportValidationErrors errors
        GoTo SafeExit
    End If

    SetStatusMessage "並び替え: ToBeファイルを再構成中"
    DebugTrace "StartSort", "ExecuteSort 呼出"
    ExecuteSort targets, mappingByToBe
    DebugTrace "StartSort", "ExecuteSort 完了"
    SetStatusMessage "並び替え: 完了"

    WriteLog LogLevelInfo(), "並び替え処理が正常終了しました。"
    MsgBox "並び替え処理が完了しました。", vbInformation + vbOKOnly

SafeExit:
    RestoreApplication
    Exit Sub

FatalError:
    WriteLog LogLevelError(), BuildUnexpectedErrorMessage("並び替え処理", Err.Number, Err.Description, Err.Source, Erl)
    MsgBox "予期しないエラーが発生しました。logシートを確認してください。", vbCritical + vbOKOnly
    RestoreApplication
End Sub

Private Sub ValidateMappingsForTargets(ByRef targets As Collection, ByRef mappingByToBe As Object, ByRef asIsToToBe As Object, ByRef errors As Collection)
    Dim requestedToBeKeys As Object
    Dim item As Variant
    Dim key As Variant
    Dim targetRow As Object
    Dim requestInfo As Object
    Dim mappingSheetName As String
    Dim toBeTable As String
    Dim toBeKey As String
    Dim tableIndex As Long
    Dim tableTotal As Long

    DebugTrace "ValidateMappingsForTargets", "開始 | targetCount=" & CStr(targets.Count)
    Set requestedToBeKeys = NewDictionary()

    For Each item In targets
        Set targetRow = item
        mappingSheetName = CStr(targetRow(KEY_TARGET_MAPPING_SHEET))
        toBeTable = CStr(targetRow(KEY_TARGET_TOBE_TABLE))
        toBeKey = BuildToBeMappingKey(mappingSheetName, toBeTable)

        If Not requestedToBeKeys.Exists(toBeKey) Then
            Set requestInfo = NewDictionary()
            requestInfo.Add KEY_TARGET_MAPPING_SHEET, mappingSheetName
            requestInfo.Add KEY_TARGET_TOBE_TABLE, toBeTable
            requestedToBeKeys.Add toBeKey, requestInfo
        End If
    Next item

    tableTotal = requestedToBeKeys.Count
    tableIndex = 0

    For Each key In requestedToBeKeys.Keys
        Set requestInfo = requestedToBeKeys(CStr(key))
        tableIndex = tableIndex + 1
        DebugTrace "ValidateMappingsForTargets", "ValidateMappingTable 呼出 | sheet=" & CStr(requestInfo(KEY_TARGET_MAPPING_SHEET)) & ", toBeTable=" & CStr(requestInfo(KEY_TARGET_TOBE_TABLE))
        SetStatusProgress "並び替え: mappingテーブルをチェック中", tableIndex, tableTotal, CStr(requestInfo(KEY_TARGET_TOBE_TABLE))
        ValidateMappingTable CStr(requestInfo(KEY_TARGET_MAPPING_SHEET)), CStr(requestInfo(KEY_TARGET_TOBE_TABLE)), mappingByToBe, asIsToToBe, errors
    Next key
End Sub

Private Sub ValidateMappingTable(ByVal mappingSheetName As String, ByVal targetToBeTable As String, ByRef mappingByToBe As Object, ByRef asIsToToBe As Object, ByRef errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim toBeTable As String
    Dim toBeColumn As String
    Dim asIsTable As String
    Dim asIsColumn As String
    Dim hasAnyValue As Boolean
    Dim duplicateToBeColumnKey As String
    Dim toBeColumnKeys As Object
    Dim toBeToAsIsTable As Object
    Dim toBeKey As String
    Dim asIsKey As String
    Dim mapRows As Collection
    Dim mapRow As Object
    Dim key As Variant
    Dim sheetToBeKeys As Object
    Dim validRowCount As Long
    Dim mappingTotalRows As Long

    DebugTrace "ValidateMappingTable", "開始 | sheet=" & mappingSheetName & ", toBeTable=" & targetToBeTable
    If Not TryGetThisWorkbookSheet(mappingSheetName, ws, errors) Then
        Exit Sub
    End If

    lastRow = GetLastRowInColumns(ws, COL_MAP_TOBE_TABLE, COL_MAP_TOBE_COLUMN, COL_MAP_ASIS_TABLE, COL_MAP_ASIS_COLUMN)
    If lastRow < START_ROW Then
        AddError errors, "[" & mappingSheetName & "]" & JpNoDataRowsSuffix()
        Exit Sub
    End If

    Set toBeColumnKeys = NewDictionary()
    Set toBeToAsIsTable = NewDictionary()
    Set sheetToBeKeys = NewDictionary()
    mappingTotalRows = lastRow - START_ROW + 1

    For rowNum = START_ROW To lastRow
        If rowNum = START_ROW Or rowNum = lastRow Or ((rowNum - START_ROW + 1) Mod PROGRESS_INTERVAL_ROWS = 0) Then
            DebugTrace "ValidateMappingTable", "走査中 | sheet=" & mappingSheetName & ", row=" & CStr(rowNum)
            SetStatusProgress "並び替え: mapping行をチェック中", rowNum - START_ROW + 1, mappingTotalRows, mappingSheetName & ":" & targetToBeTable
        End If
        toBeTable = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_TABLE).Value)
        toBeColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_COLUMN).Value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_TABLE).Value)
        asIsColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_COLUMN).Value)

        hasAnyValue = (Len(toBeTable) > 0 Or Len(toBeColumn) > 0 Or Len(asIsTable) > 0 Or Len(asIsColumn) > 0)
        If Not hasAnyValue Then
            GoTo NextMappingRow
        End If

        If NormalizeKey(toBeTable) <> NormalizeKey(targetToBeTable) Then
            GoTo NextMappingRow
        End If

        If Len(toBeTable) = 0 Then
            AddError errors, mappingSheetName & "!B" & rowNum & JpRequiredSuffix()
            GoTo NextMappingRow
        End If

        If Len(toBeColumn) = 0 Then
            AddError errors, mappingSheetName & "!D" & rowNum & JpRequiredSuffix()
            GoTo NextMappingRow
        End If

        If (Len(asIsTable) = 0 Xor Len(asIsColumn) = 0) Then
            AddError errors, JpMappingAsIsPairError(mappingSheetName, rowNum)
            GoTo NextMappingRow
        End If

        toBeKey = BuildToBeMappingKey(mappingSheetName, toBeTable)
        duplicateToBeColumnKey = toBeKey & "|" & NormalizeKey(toBeColumn)
        If toBeColumnKeys.Exists(duplicateToBeColumnKey) Then
            AddError errors, JpDuplicateToBeColumnError(mappingSheetName, rowNum, toBeTable, toBeColumn)
            GoTo NextMappingRow
        End If
        toBeColumnKeys.Add duplicateToBeColumnKey, rowNum

        If Not mappingByToBe.Exists(toBeKey) Then
            Set mapRows = New Collection
            mappingByToBe.Add toBeKey, mapRows
        End If

        If Not sheetToBeKeys.Exists(toBeKey) Then
            sheetToBeKeys.Add toBeKey, True
        End If

        Set mapRow = NewDictionary()
        mapRow.Add KEY_MAP_ROWNUM, rowNum
        mapRow.Add KEY_MAP_TOBE_TABLE, toBeTable
        mapRow.Add KEY_MAP_TOBE_COLUMN, toBeColumn
        mapRow.Add KEY_MAP_ASIS_TABLE, asIsTable
        mapRow.Add KEY_MAP_ASIS_COLUMN, asIsColumn

        Set mapRows = mappingByToBe(toBeKey)
        mapRows.Add mapRow
        validRowCount = validRowCount + 1

        If Len(asIsTable) > 0 Then
            asIsKey = BuildAsIsMappingKey(mappingSheetName, asIsTable)

            If asIsToToBe.Exists(asIsKey) Then
                If CStr(asIsToToBe(asIsKey)) <> toBeKey Then
                    AddError errors, JpAsIsToToBeConflictError(mappingSheetName, rowNum, asIsTable)
                End If
            Else
                asIsToToBe.Add asIsKey, toBeKey
            End If

            If toBeToAsIsTable.Exists(toBeKey) Then
                If CStr(toBeToAsIsTable(toBeKey)) <> asIsKey Then
                    AddError errors, JpToBeToAsIsConflictError(mappingSheetName, toBeTable)
                End If
            Else
                toBeToAsIsTable.Add toBeKey, asIsKey
            End If
        End If

NextMappingRow:
    Next rowNum

    If validRowCount = 0 Then Exit Sub

    For Each key In sheetToBeKeys.Keys
        If Not toBeToAsIsTable.Exists(CStr(key)) Then
            Set mapRows = mappingByToBe(CStr(key))
            Set mapRow = mapRows(1)
            AddError errors, JpToBeNoAsIsError(mappingSheetName, CStr(mapRow(KEY_MAP_TOBE_TABLE)))
        End If
    Next key
End Sub

Private Sub ValidateTarget(ByRef targets As Collection, ByRef errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim targetLine As TargetLineData
    Dim targetRow As Object
    Dim mappingSheetHasToBe As Object
    Dim totalRows As Long

    DebugTrace "ValidateTarget", "開始"
    If Not TryGetThisWorkbookSheet(SHEET_TARGET, ws, errors) Then
        Exit Sub
    End If

    Set mappingSheetHasToBe = NewDictionary()
    lastRow = GetTargetLastRow(ws)
    If lastRow < START_ROW Then
        AddError errors, "target" & JpNoDataRowsSuffix()
        Exit Sub
    End If

    totalRows = lastRow - START_ROW + 1

    For rowNum = START_ROW To lastRow
        If rowNum = START_ROW Or rowNum = lastRow Or ((rowNum - START_ROW + 1) Mod PROGRESS_INTERVAL_ROWS = 0) Then
            DebugTrace "ValidateTarget", "走査中 | row=" & CStr(rowNum)
            SetStatusProgress "並び替え: target行をチェック中", rowNum - START_ROW + 1, totalRows
        End If
        If Not ReadTargetLine(ws, rowNum, targetLine) Then
            GoTo NextTargetRow
        End If

        If Len(targetLine.MappingSheetName) = 0 Then
            GoTo NextTargetRow
        End If

        If Not HasAnyToBeTableInMappingSheet(targetLine.MappingSheetName, mappingSheetHasToBe) Then
            GoTo NextTargetRow
        End If

        If Not ValidateTargetLine(rowNum, targetLine, errors) Then
            GoTo NextTargetRow
        End If

        Set targetRow = CreateTargetRow(rowNum, targetLine)

        targets.Add targetRow

NextTargetRow:
    Next rowNum

    If targets.Count = 0 Then
        AddError errors, "target" & JpNoValidRowsSuffix()
    End If
End Sub

Private Function GetTargetLastRow(ByVal ws As Worksheet) As Long
    GetTargetLastRow = GetLastRowInColumns6(ws, COL_TARGET_MAPPING_SHEET, COL_TARGET_FILE, COL_TARGET_SHEET, COL_TARGET_TOBE_TABLE, COL_TARGET_HEADER_CELL, COL_TARGET_DATA_CELL)
End Function

Private Function ReadTargetLine(ByVal ws As Worksheet, ByVal rowNum As Long, ByRef targetLine As TargetLineData) As Boolean
    targetLine.MappingSheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_MAPPING_SHEET).Value)
    targetLine.FilePath = TrimSafe(ws.Cells(rowNum, COL_TARGET_FILE).Value)
    targetLine.SheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_SHEET).Value)
    targetLine.ToBeTable = TrimSafe(ws.Cells(rowNum, COL_TARGET_TOBE_TABLE).Value)
    targetLine.HeaderCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_HEADER_CELL).Value)
    targetLine.DataCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_DATA_CELL).Value)

    ReadTargetLine = HasAnyTargetInput(targetLine)
End Function

Private Function HasAnyTargetInput(ByRef targetLine As TargetLineData) As Boolean
    HasAnyTargetInput = (Len(targetLine.MappingSheetName) > 0 _
        Or Len(targetLine.FilePath) > 0 _
        Or Len(targetLine.SheetName) > 0 _
        Or Len(targetLine.ToBeTable) > 0 _
        Or Len(targetLine.HeaderCell) > 0 _
        Or Len(targetLine.DataCell) > 0)
End Function

Private Function HasAnyToBeTableInMappingSheet(ByVal mappingSheetName As String, ByRef cache As Object) As Boolean
    Dim key As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long

    key = NormalizeKey(mappingSheetName)
    If cache.Exists(key) Then
        HasAnyToBeTableInMappingSheet = CBool(cache(key))
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(mappingSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        cache.Add key, False
        HasAnyToBeTableInMappingSheet = False
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_MAP_TOBE_TABLE).End(xlUp).Row
    If lastRow < START_ROW Then
        cache.Add key, False
        HasAnyToBeTableInMappingSheet = False
        Exit Function
    End If

    For rowNum = START_ROW To lastRow
        If Len(TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_TABLE).Value)) > 0 Then
            cache.Add key, True
            HasAnyToBeTableInMappingSheet = True
            Exit Function
        End If
    Next rowNum

    cache.Add key, False
    HasAnyToBeTableInMappingSheet = False
End Function

Private Function ValidateTargetLine(ByVal rowNum As Long, ByRef targetLine As TargetLineData, ByRef errors As Collection) As Boolean
    Dim headerRow As Long
    Dim headerCol As Long
    Dim dataRow As Long
    Dim dataCol As Long
    Dim rowHasError As Boolean

    rowHasError = False

    If Len(targetLine.MappingSheetName) = 0 Then
        AddError errors, "target!B" & rowNum & JpRequiredSuffix()
        rowHasError = True
    End If

    If Len(targetLine.FilePath) = 0 Then
        AddError errors, "target!C" & rowNum & JpRequiredSuffix()
        ValidateTargetLine = False
        Exit Function
    End If

    If Not IsAbsolutePath(targetLine.FilePath) Then
        AddError errors, "target!C" & rowNum & JpAbsolutePathSuffix() & targetLine.FilePath
        rowHasError = True
    End If

    If Len(targetLine.SheetName) = 0 Then
        AddError errors, "target!D" & rowNum & JpRequiredSuffix()
        rowHasError = True
    End If

    If Len(targetLine.ToBeTable) = 0 Then
        AddError errors, "target!E" & rowNum & JpRequiredSuffix()
        rowHasError = True
    End If

    If Not ParseA1Address(targetLine.HeaderCell, headerCol, headerRow) Then
        AddError errors, "target!F" & rowNum & JpInvalidA1Suffix() & targetLine.HeaderCell
        rowHasError = True
    End If

    If Not ParseA1Address(targetLine.DataCell, dataCol, dataRow) Then
        AddError errors, "target!G" & rowNum & JpInvalidA1Suffix() & targetLine.DataCell
        rowHasError = True
    End If

    If ParseA1Address(targetLine.HeaderCell, headerCol, headerRow) And ParseA1Address(targetLine.DataCell, dataCol, dataRow) Then
        If dataRow <= headerRow Then
            AddError errors, TargetRowLabel(rowNum) & JpTargetDataStartRowOrderSuffix()
            rowHasError = True
        End If
    End If

    ValidateTargetLine = Not rowHasError
End Function

Private Function CreateTargetRow(ByVal rowNum As Long, ByRef targetLine As TargetLineData) As Object
    Dim targetRow As Object

    Set targetRow = NewDictionary()
    targetRow.Add KEY_TARGET_ROWNUM, rowNum
    targetRow.Add KEY_TARGET_MAPPING_SHEET, targetLine.MappingSheetName
    targetRow.Add KEY_TARGET_FILE_PATH, targetLine.FilePath
    targetRow.Add KEY_TARGET_SHEET_NAME, targetLine.SheetName
    targetRow.Add KEY_TARGET_TOBE_TABLE, targetLine.ToBeTable
    targetRow.Add KEY_TARGET_HEADER_CELL, UCase$(targetLine.HeaderCell)
    targetRow.Add KEY_TARGET_DATA_CELL, UCase$(targetLine.DataCell)
    targetRow.Add KEY_TARGET_OUTPUT_FILE_PATH, BuildToBeFilePath(targetLine.FilePath)

    Set CreateTargetRow = targetRow
End Function

Private Sub BindTargetsToMapping(ByRef targets As Collection, ByRef mappingByToBe As Object, ByRef errors As Collection)
    Dim item As Variant
    Dim targetRow As Object
    Dim toBeKey As String
    Dim mappingSheetName As String

    DebugTrace "BindTargetsToMapping", "開始 | targetCount=" & CStr(targets.Count)
    For Each item In targets
        Set targetRow = item
        mappingSheetName = CStr(targetRow(KEY_TARGET_MAPPING_SHEET))
        toBeKey = BuildToBeMappingKey(mappingSheetName, CStr(targetRow(KEY_TARGET_TOBE_TABLE)))
        DebugTrace "BindTargetsToMapping", "紐づけ | row=" & CStr(targetRow(KEY_TARGET_ROWNUM)) & ", key=" & toBeKey

        If Not mappingByToBe.Exists(toBeKey) Then
            AddError errors, JpTargetToBeTableMappingNotFoundError(CLng(targetRow(KEY_TARGET_ROWNUM)), CStr(targetRow(KEY_TARGET_TOBE_TABLE)), mappingSheetName)
        Else
            targetRow.Add KEY_TARGET_TOBE_TABLE_KEY, toBeKey
        End If
    Next item
End Sub

Private Sub ValidateFilesSheetsAndColumns(ByRef targets As Collection, ByRef mappingByToBe As Object, ByRef errors As Collection)
    Dim targetItem As Variant
    Dim targetRow As Object
    Dim filePath As String
    Dim outputFilePath As String
    Dim sheetName As String
    Dim mappingSheetName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedBooks As Object
    Dim toBeKey As String
    Dim mappingRows As Collection
    Dim mapItem As Variant
    Dim mapRow As Object
    Dim asIsColumn As String
    Dim headerLookup As Object
    Dim rowNum As Long
    Dim targetCount As Long
    Dim targetIndex As Long

    DebugTrace "ValidateFilesSheetsAndColumns", "開始 | targetCount=" & CStr(targets.Count)
    Set openedBooks = NewDictionary()

    On Error GoTo Cleanup

    targetCount = targets.Count
    targetIndex = 0

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow(KEY_TARGET_FILE_PATH))
        outputFilePath = CStr(targetRow(KEY_TARGET_OUTPUT_FILE_PATH))
        rowNum = CLng(targetRow(KEY_TARGET_ROWNUM))
        targetIndex = targetIndex + 1
        DebugTrace "ValidateFilesSheetsAndColumns", "出力ファイル事前チェック | row=" & CStr(rowNum) & ", file=" & filePath & ", output=" & outputFilePath
        SetStatusProgress "並び替え: 出力ファイルを事前チェック中", targetIndex, targetCount

        If Not FileExists(filePath) Then
            AddError errors, TargetRowLabel(rowNum) & " のファイルが存在しません: " & filePath
        End If

        If FileExists(outputFilePath) Then
            AddError errors, TargetRowLabel(rowNum) & " の出力先ファイルが既に存在します: " & outputFilePath
        End If
    Next targetItem

    targetIndex = 0

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow(KEY_TARGET_FILE_PATH))
        sheetName = CStr(targetRow(KEY_TARGET_SHEET_NAME))
        mappingSheetName = CStr(targetRow(KEY_TARGET_MAPPING_SHEET))
        rowNum = CLng(targetRow(KEY_TARGET_ROWNUM))
        targetIndex = targetIndex + 1
        DebugTrace "ValidateFilesSheetsAndColumns", "元シート/カラム事前チェック | row=" & CStr(rowNum) & ", sheet=" & sheetName
        SetStatusProgress "並び替え: 元シート/カラムを事前チェック中", targetIndex, targetCount

        If Not FileExists(filePath) Then
            GoTo NextTarget
        End If

        Set wb = GetOrOpenWorkbook(filePath, True, openedBooks, errors, rowNum)
        If wb Is Nothing Then
            GoTo NextTarget
        End If

        Set ws = Nothing
        If Not TryGetWorksheetFromWorkbook(wb, sheetName, ws) Then
            AddError errors, TargetRowLabel(rowNum) & " のシートが存在しません: " & filePath & " / " & sheetName
            GoTo NextTarget
        End If

        toBeKey = ""
        If targetRow.Exists(KEY_TARGET_TOBE_TABLE_KEY) Then
            toBeKey = CStr(targetRow(KEY_TARGET_TOBE_TABLE_KEY))
        End If

        If Len(toBeKey) = 0 Then
            GoTo NextTarget
        End If

        If Not mappingByToBe.Exists(toBeKey) Then
            AddError errors, JpTargetToBeMappingNotFoundError(rowNum, mappingSheetName)
            GoTo NextTarget
        End If

        Set headerLookup = BuildHeaderLookupForValidation(ws, CStr(targetRow(KEY_TARGET_HEADER_CELL)), errors, rowNum, filePath, sheetName)
        Set mappingRows = mappingByToBe(toBeKey)
        DebugTrace "ValidateFilesSheetsAndColumns", "ヘッダ確認 | row=" & CStr(rowNum) & ", mapCount=" & CStr(mappingRows.Count)

        For Each mapItem In mappingRows
            Set mapRow = mapItem
            asIsColumn = CStr(mapRow(KEY_MAP_ASIS_COLUMN))

            If Len(asIsColumn) > 0 Then
                If Not headerLookup.Exists(NormalizeKey(asIsColumn)) Then
                    AddError errors, TargetRowLabel(rowNum) & " で必要なAsIsカラムが見つかりません: " & asIsColumn & " (" & filePath & " / " & sheetName & ")"
                End If
            End If
        Next mapItem

NextTarget:
    Next targetItem

Cleanup:
    CloseWorkbooks openedBooks, False
End Sub

Private Sub ExecuteSort(ByRef targets As Collection, ByRef mappingByToBe As Object)
    Dim openedBooks As Object
    Dim outputPlans As Object
    Dim outputPlan As Object
    Dim targetSheets As Object
    Dim targetItem As Variant
    Dim targetRow As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim mappingRows As Collection
    Dim sourceFilePath As String
    Dim outputFilePath As String
    Dim sheetName As String
    Dim toBeKey As String
    Dim rowNum As Long
    Dim key As Variant
    Dim planKey As String
    Dim targetCount As Long
    Dim targetIndex As Long
    Dim planCount As Long
    Dim planIndex As Long
    Dim saveCount As Long
    Dim saveIndex As Long

    DebugTrace "ExecuteSort", "開始 | targetCount=" & CStr(targets.Count)
    Set openedBooks = NewDictionary()
    Set outputPlans = NewDictionary()

    On Error GoTo ExecError

    targetCount = targets.Count
    targetIndex = 0

    For Each targetItem In targets
        Set targetRow = targetItem
        sourceFilePath = CStr(targetRow(KEY_TARGET_FILE_PATH))
        outputFilePath = CStr(targetRow(KEY_TARGET_OUTPUT_FILE_PATH))
        sheetName = CStr(targetRow(KEY_TARGET_SHEET_NAME))
        planKey = NormalizeKey(sourceFilePath)
        targetIndex = targetIndex + 1
        DebugTrace "ExecuteSort", "出力計画作成 | source=" & sourceFilePath & ", output=" & outputFilePath & ", sheet=" & sheetName
        SetStatusProgress "並び替え: 出力ファイル計画を作成中", targetIndex, targetCount

        If Not outputPlans.Exists(planKey) Then
            Set outputPlan = NewDictionary()
            outputPlan.Add KEY_PLAN_SOURCE_FILE, sourceFilePath
            outputPlan.Add KEY_PLAN_OUTPUT_FILE, outputFilePath
            Set targetSheets = NewDictionary()
            outputPlan.Add KEY_PLAN_TARGET_SHEETS, targetSheets
            outputPlans.Add planKey, outputPlan
        End If

        Set outputPlan = outputPlans(planKey)
        Set targetSheets = outputPlan(KEY_PLAN_TARGET_SHEETS)
        If Not targetSheets.Exists(NormalizeKey(sheetName)) Then
            targetSheets.Add NormalizeKey(sheetName), sheetName
        End If
    Next targetItem

    planCount = outputPlans.Count
    planIndex = 0

    For Each key In outputPlans.Keys
        Set outputPlan = outputPlans(CStr(key))
        sourceFilePath = CStr(outputPlan(KEY_PLAN_SOURCE_FILE))
        outputFilePath = CStr(outputPlan(KEY_PLAN_OUTPUT_FILE))
        planIndex = planIndex + 1
        DebugTrace "ExecuteSort", "TOBE作成 | source=" & sourceFilePath & ", output=" & outputFilePath
        SetStatusProgress "並び替え: _TOBEファイルを作成中", planIndex, planCount

        FileCopy sourceFilePath, outputFilePath
        WriteLog LogLevelInfo(), "ファイルをコピーしました: " & outputFilePath

        Set wb = GetOrOpenWorkbookForRun(outputFilePath, openedBooks)
        Set targetSheets = outputPlan(KEY_PLAN_TARGET_SHEETS)
        KeepOnlyTargetSheets wb, targetSheets
        ResetTargetSheetView wb, targetSheets
    Next key

    targetIndex = 0

    For Each targetItem In targets
        Set targetRow = targetItem
        rowNum = CLng(targetRow(KEY_TARGET_ROWNUM))
        outputFilePath = CStr(targetRow(KEY_TARGET_OUTPUT_FILE_PATH))
        sheetName = CStr(targetRow(KEY_TARGET_SHEET_NAME))
        toBeKey = CStr(targetRow(KEY_TARGET_TOBE_TABLE_KEY))
        targetIndex = targetIndex + 1
        DebugTrace "ExecuteSort", "シート再構成 | row=" & CStr(rowNum) & ", output=" & outputFilePath & ", sheet=" & sheetName & ", key=" & toBeKey
        SetStatusProgress "並び替え: シートを再構成中", targetIndex, targetCount, sheetName

        Set wb = GetOrOpenWorkbookForRun(outputFilePath, openedBooks)
        Set ws = wb.Worksheets(sheetName)

        Set mappingRows = mappingByToBe(toBeKey)
        RebuildToBeSheet ws, CStr(targetRow(KEY_TARGET_HEADER_CELL)), CStr(targetRow(KEY_TARGET_DATA_CELL)), mappingRows

        WriteLog LogLevelInfo(), TargetRowLabel(rowNum) & " を処理しました: " & outputFilePath & " / " & sheetName
    Next targetItem

    saveCount = openedBooks.Count
    saveIndex = 0

    For Each key In openedBooks.Keys
        saveIndex = saveIndex + 1
        DebugTrace "ExecuteSort", "保存 | file=" & CStr(key)
        SetStatusProgress "並び替え: ファイルを保存中", saveIndex, saveCount
        openedBooks(key).Save
    Next key

    CloseWorkbooks openedBooks, False
    Exit Sub

ExecError:
    DebugTrace "ExecuteSort", "エラー捕捉 | err=" & CStr(Err.Number) & ", source=" & Err.Source
    CloseWorkbooks openedBooks, False
    If Len(Err.Source) > 0 Then
        Err.Raise Err.Number, "ExecuteSort:" & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "ExecuteSort", Err.Description
    End If
End Sub

Private Sub KeepOnlyTargetSheets(ByVal wb As Workbook, ByVal targetSheets As Object)
    Dim ws As Worksheet
    Dim sh As Object
    Dim keepSheet As Worksheet
    Dim deleteNames As Collection
    Dim item As Variant

    Set deleteNames = New Collection

    For Each ws In wb.Worksheets
        If targetSheets.Exists(NormalizeKey(ws.Name)) Then
            If keepSheet Is Nothing Then
                Set keepSheet = ws
            End If
        End If
    Next ws

    For Each sh In wb.Sheets
        If Not targetSheets.Exists(NormalizeKey(CStr(sh.Name))) Then
            deleteNames.Add CStr(sh.Name)
        End If
    Next sh

    If keepSheet Is Nothing Then
        Err.Raise vbObjectError + 3001, "KeepOnlyTargetSheets", "targetで指定されたシートが_TOBE内に見つかりません。"
    End If

    keepSheet.Visible = xlSheetVisible

    For Each item In deleteNames
        wb.Sheets(CStr(item)).Delete
    Next item
End Sub

Private Sub ResetTargetSheetView(ByVal wb As Workbook, ByVal targetSheets As Object)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If targetSheets.Exists(NormalizeKey(ws.Name)) Then
            ClearSheetFilters ws
            ws.Cells.EntireColumn.Hidden = False
        End If
    Next ws
End Sub

Private Sub ClearSheetFilters(ByVal ws As Worksheet)
    Dim listObj As ListObject

    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    For Each listObj In ws.ListObjects
        On Error Resume Next
        listObj.AutoFilter.ShowAllData
        On Error GoTo 0
    Next listObj
End Sub

Private Sub RebuildToBeSheet(ByVal ws As Worksheet, ByVal headerCellAddress As String, ByVal dataCellAddress As String, ByVal mappingRows As Collection)
    Dim headerLookup As Object
    Dim dataStartRow As Long
    Dim dataStartCol As Long
    Dim headerRow As Long
    Dim headerStartCol As Long
    Dim rowCount As Long
    Dim mapCount As Long
    Dim outputArray() As Variant
    Dim mapIndex As Long
    Dim dataIndex As Long
    Dim mapItem As Variant
    Dim mapRow As Object
    Dim toBeColumn As String
    Dim asIsColumn As String
    Dim sourceCol As Long
    Dim columnData As Variant
    Dim lastUsedRow As Long
    Dim lastUsedCol As Long

    On Error GoTo RebuildError
    DebugTrace "RebuildToBeSheet", "開始 | sheet=" & ws.Name & ", header=" & headerCellAddress & ", data=" & dataCellAddress & ", mapCount=" & CStr(mappingRows.Count)
    Set headerLookup = BuildHeaderLookupStrict(ws, headerCellAddress)

    ParseA1Address dataCellAddress, dataStartCol, dataStartRow
    ParseA1Address headerCellAddress, headerStartCol, headerRow
    DebugTrace "RebuildToBeSheet", "A1解析 | dataRow=" & CStr(dataStartRow) & ", dataCol=" & CStr(dataStartCol) & ", headerRow=" & CStr(headerRow) & ", headerCol=" & CStr(headerStartCol)

    lastUsedRow = GetLastUsedRow(ws)
    If lastUsedRow >= dataStartRow Then
        rowCount = lastUsedRow - dataStartRow + 1
    Else
        rowCount = 0
    End If

    mapCount = mappingRows.Count
    ReDim outputArray(1 To rowCount + 1, 1 To mapCount)

    mapIndex = 0
    For Each mapItem In mappingRows
        mapIndex = mapIndex + 1
        Set mapRow = mapItem

        toBeColumn = CStr(mapRow(KEY_MAP_TOBE_COLUMN))
        asIsColumn = CStr(mapRow(KEY_MAP_ASIS_COLUMN))

        outputArray(1, mapIndex) = toBeColumn

        If Len(asIsColumn) > 0 And rowCount > 0 Then
            sourceCol = CLng(headerLookup(NormalizeKey(asIsColumn)))
            columnData = ws.Range(ws.Cells(dataStartRow, sourceCol), ws.Cells(dataStartRow + rowCount - 1, sourceCol)).Value2

            If IsArray(columnData) Then
                For dataIndex = 1 To rowCount
                    outputArray(dataIndex + 1, mapIndex) = columnData(dataIndex, 1)
                Next dataIndex
            Else
                ' 1セル取得時は配列ではなく単一値になるため直接設定する
                outputArray(2, mapIndex) = columnData
                DebugTrace "RebuildToBeSheet", "単一セル値を設定 | column=" & asIsColumn
            End If
        End If
    Next mapItem

    lastUsedRow = GetLastUsedRow(ws)
    lastUsedCol = GetLastUsedCol(ws)

    If lastUsedRow < headerRow Then
        lastUsedRow = headerRow
    End If
    If lastUsedCol < headerStartCol Then
        lastUsedCol = headerStartCol
    End If

    DebugTrace "RebuildToBeSheet", "出力書込 | rowCount=" & CStr(rowCount) & ", mapCount=" & CStr(mapCount)
    ws.Range(ws.Cells(headerRow, headerStartCol), ws.Cells(lastUsedRow, lastUsedCol)).ClearContents
    ws.Range(ws.Cells(headerRow, headerStartCol), ws.Cells(headerRow + rowCount, headerStartCol + mapCount - 1)).Value2 = outputArray
    Exit Sub

RebuildError:
    DebugTrace "RebuildToBeSheet", "エラー捕捉 | err=" & CStr(Err.Number) & ", source=" & Err.Source & ", sheet=" & ws.Name
    If Len(Err.Source) > 0 Then
        Err.Raise Err.Number, "RebuildToBeSheet:" & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "RebuildToBeSheet", Err.Description
    End If
End Sub

Private Function BuildHeaderLookupForValidation(ByVal ws As Worksheet, ByVal headerCellAddress As String, ByRef errors As Collection, ByVal targetRowNum As Long, ByVal filePath As String, ByVal sheetName As String) As Object
    Dim headerLookup As Object
    Dim duplicateHeaders As Collection
    Dim duplicateItem As Variant

    Set headerLookup = BuildHeaderLookupCore(ws, headerCellAddress, duplicateHeaders)

    For Each duplicateItem In duplicateHeaders
        AddError errors, "target 行" & targetRowNum & " のヘッダに同名カラムが重複しています: " & CStr(duplicateItem) & " (" & filePath & " / " & sheetName & ")"
    Next duplicateItem

    Set BuildHeaderLookupForValidation = headerLookup
End Function

Private Function BuildHeaderLookupStrict(ByVal ws As Worksheet, ByVal headerCellAddress As String) As Object
    Dim headerLookup As Object
    Dim duplicateHeaders As Collection

    Set headerLookup = BuildHeaderLookupCore(ws, headerCellAddress, duplicateHeaders)

    If duplicateHeaders.Count > 0 Then
        Err.Raise vbObjectError + 2001, "BuildHeaderLookupStrict", "ヘッダに同名カラムが重複しています: " & CStr(duplicateHeaders(1))
    End If

    Set BuildHeaderLookupStrict = headerLookup
End Function

Private Function BuildHeaderLookupCore(ByVal ws As Worksheet, ByVal headerCellAddress As String, ByRef duplicateHeaders As Collection) As Object
    Dim headerLookup As Object
    Dim headerRow As Long
    Dim startCol As Long
    Dim lastCol As Long
    Dim col As Long
    Dim headerName As String
    Dim key As String

    Set headerLookup = NewDictionary()
    Set duplicateHeaders = New Collection

    ParseA1Address headerCellAddress, startCol, headerRow

    If Application.WorksheetFunction.CountA(ws.Rows(headerRow)) = 0 Then
        Set BuildHeaderLookupCore = headerLookup
        Exit Function
    End If

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < startCol Then
        Set BuildHeaderLookupCore = headerLookup
        Exit Function
    End If

    For col = startCol To lastCol
        headerName = TrimSafe(ws.Cells(headerRow, col).value)
        If Len(headerName) > 0 Then
            key = NormalizeKey(headerName)
            If headerLookup.Exists(key) Then
                duplicateHeaders.Add headerName
            Else
                headerLookup.Add key, col
            End If
        End If
    Next col

    Set BuildHeaderLookupCore = headerLookup
End Function

Private Function BuildToBeMappingKey(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    BuildToBeMappingKey = NormalizeKey(mappingSheetName) & "|" & NormalizeKey(toBeTable)
End Function

Private Function BuildAsIsMappingKey(ByVal mappingSheetName As String, ByVal asIsTable As String) As String
    BuildAsIsMappingKey = NormalizeKey(mappingSheetName) & "|" & NormalizeKey(asIsTable)
End Function
