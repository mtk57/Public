Attribute VB_Name = "modAsIsToToBeSorter"
Option Explicit

Private Const VER = "0.5.0"

Private Const SHEET_MAPPING As String = "mapping"
Private Const SHEET_TARGET As String = "target"
Private Const SHEET_MACRO As String = "log"

Private Const TOBE_FILE_SUFFIX As String = "_TOBE"
Private Const START_ROW As Long = 5

' mapping sheet columns
Private Const COL_MAP_TOBE_TABLE As Long = 2   ' B
Private Const COL_MAP_TOBE_COLUMN As Long = 4  ' D
Private Const COL_MAP_ASIS_TABLE As Long = 15  ' O
Private Const COL_MAP_ASIS_COLUMN As Long = 17 ' Q

' target sheet columns
Private Const COL_TARGET_FILE As Long = 3       ' C
Private Const COL_TARGET_SHEET As Long = 4      ' D
Private Const COL_TARGET_ASIS_TABLE As Long = 5 ' E
Private Const COL_TARGET_HEADER_CELL As Long = 6 ' F
Private Const COL_TARGET_DATA_CELL As Long = 7   ' G

' log area (B:D)
Private Const LOG_ROW_HEADER As Long = 1
Private Const LOG_ROW_START As Long = 2
Private Const LOG_COL_TIME As Long = 2
Private Const LOG_COL_LEVEL As Long = 3
Private Const LOG_COL_MESSAGE As Long = 4

Private Const MAX_EXCEL_ROW As Long = 1048576
Private Const MAX_EXCEL_COL As Long = 16384

Private gLogNextRow As Long
Private gPrevScreenUpdating As Boolean
Private gPrevEnableEvents As Boolean
Private gPrevDisplayAlerts As Boolean
Private gPrevCalculation As XlCalculation

Public Sub 並び替え開始()
    StartSort
End Sub

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

    Set errors = New Collection
    Set mappingByToBe = NewDictionary()
    Set asIsToToBe = NewDictionary()
    Set targets = New Collection

    ValidateMapping mappingByToBe, asIsToToBe, errors
    ValidateTarget targets, errors
    BindTargetsToMapping targets, asIsToToBe, errors
    ValidateFilesSheetsAndColumns targets, mappingByToBe, errors

    If errors.Count > 0 Then
        ReportValidationErrors errors
        GoTo SafeExit
    End If

    ExecuteSort targets, mappingByToBe

    WriteLog LogLevelInfo(), "並び替え処理が正常終了しました。"
    MsgBox "並び替え処理が完了しました。", vbInformation + vbOKOnly

SafeExit:
    RestoreApplication
    Exit Sub

FatalError:
    WriteLog LogLevelError(), "予期しないエラー: (" & Err.Number & ") " & Err.Description
    MsgBox ChrW(&H4E88) & ChrW(&H671F) & ChrW(&H3057) & ChrW(&H306A) & ChrW(&H3044) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H304C) & ChrW(&H767A) & ChrW(&H751F) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002) & ChrW(&H6C) & ChrW(&H6F) & ChrW(&H67) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002), vbCritical + vbOKOnly
    RestoreApplication
End Sub

Private Sub ValidateMapping(ByRef mappingByToBe As Object, ByRef asIsToToBe As Object, ByRef errors As Collection)
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

    If Not TryGetThisWorkbookSheet(SHEET_MAPPING, ws, errors) Then
        Exit Sub
    End If

    lastRow = GetLastRowInColumns(ws, COL_MAP_TOBE_TABLE, COL_MAP_TOBE_COLUMN, COL_MAP_ASIS_TABLE, COL_MAP_ASIS_COLUMN)
    If lastRow < START_ROW Then
        AddError errors, "mappingシートにデータ行がありません。"
        Exit Sub
    End If

    Set toBeColumnKeys = NewDictionary()
    Set toBeToAsIsTable = NewDictionary()

    For rowNum = START_ROW To lastRow
        toBeTable = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_TABLE).value)
        toBeColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_COLUMN).value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_TABLE).value)
        asIsColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_COLUMN).value)

        hasAnyValue = (Len(toBeTable) > 0 Or Len(toBeColumn) > 0 Or Len(asIsTable) > 0 Or Len(asIsColumn) > 0)
        If Not hasAnyValue Then
            GoTo NextMappingRow
        End If

        If Len(toBeTable) = 0 Then
            AddError errors, "mapping!B" & rowNum & " が空です。ToBeテーブル名(物理名)は必須です。"
            GoTo NextMappingRow
        End If

        If Len(toBeColumn) = 0 Then
            AddError errors, "mapping!D" & rowNum & " が空です。ToBeカラム名(物理名)は必須です。"
            GoTo NextMappingRow
        End If

        If (Len(asIsTable) = 0 Xor Len(asIsColumn) = 0) Then
            AddError errors, "mapping 行" & rowNum & " の AsIsテーブル名(O列)とAsIsカラム名(Q列)は両方入力か両方空欄にしてください。"
            GoTo NextMappingRow
        End If

        toBeKey = NormalizeKey(toBeTable)
        duplicateToBeColumnKey = toBeKey & "|" & NormalizeKey(toBeColumn)
        If toBeColumnKeys.Exists(duplicateToBeColumnKey) Then
            AddError errors, "mapping 行" & rowNum & " のToBeカラムが重複しています（同一ToBeテーブル内で重複不可）: " & toBeTable & "." & toBeColumn
            GoTo NextMappingRow
        End If
        toBeColumnKeys.Add duplicateToBeColumnKey, rowNum

        If Not mappingByToBe.Exists(toBeKey) Then
            Set mapRows = New Collection
            mappingByToBe.Add toBeKey, mapRows
        End If

        Set mapRow = NewDictionary()
        mapRow.Add "RowNum", rowNum
        mapRow.Add "ToBeTable", toBeTable
        mapRow.Add "ToBeColumn", toBeColumn
        mapRow.Add "AsIsTable", asIsTable
        mapRow.Add "AsIsColumn", asIsColumn

        Set mapRows = mappingByToBe(toBeKey)
        mapRows.Add mapRow

        If Len(asIsTable) > 0 Then
            asIsKey = NormalizeKey(asIsTable)

            If asIsToToBe.Exists(asIsKey) Then
                If CStr(asIsToToBe(asIsKey)) <> toBeKey Then
                    AddError errors, "mapping 行" & rowNum & " でAsIsテーブル " & asIsTable & " が複数ToBeテーブルに紐づいています。"
                End If
            Else
                asIsToToBe.Add asIsKey, toBeKey
            End If

            If toBeToAsIsTable.Exists(toBeKey) Then
                If CStr(toBeToAsIsTable(toBeKey)) <> asIsKey Then
                    AddError errors, "mapping のToBeテーブル " & toBeTable & " に複数のAsIsテーブル名が定義されています。"
                End If
            Else
                toBeToAsIsTable.Add toBeKey, asIsKey
            End If
        End If

NextMappingRow:
    Next rowNum

    If mappingByToBe.Count = 0 Then
        AddError errors, "mappingシートに有効なマッピング定義がありません。"
        Exit Sub
    End If

    For Each key In mappingByToBe.Keys
        If Not toBeToAsIsTable.Exists(CStr(key)) Then
            Set mapRows = mappingByToBe(CStr(key))
            Set mapRow = mapRows(1)
            AddError errors, "mapping のToBeテーブル " & CStr(mapRow("ToBeTable")) & " にはAsIsテーブルが1件も定義されていません。"
        End If
    Next key
End Sub

Private Sub ValidateTarget(ByRef targets As Collection, ByRef errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim filePath As String
    Dim sheetName As String
    Dim asIsTable As String
    Dim headerCell As String
    Dim dataCell As String
    Dim hasAnyValue As Boolean
    Dim headerRow As Long
    Dim headerCol As Long
    Dim dataRow As Long
    Dim dataCol As Long
    Dim targetRow As Object
    Dim rowHasError As Boolean

    If Not TryGetThisWorkbookSheet(SHEET_TARGET, ws, errors) Then
        Exit Sub
    End If

    lastRow = GetLastRowInColumns5(ws, COL_TARGET_FILE, COL_TARGET_SHEET, COL_TARGET_ASIS_TABLE, COL_TARGET_HEADER_CELL, COL_TARGET_DATA_CELL)
    If lastRow < START_ROW Then
        AddError errors, "targetシートにデータ行がありません。"
        Exit Sub
    End If

    For rowNum = START_ROW To lastRow
        rowHasError = False
        filePath = TrimSafe(ws.Cells(rowNum, COL_TARGET_FILE).value)
        sheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_SHEET).value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_TARGET_ASIS_TABLE).value)
        headerCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_HEADER_CELL).value)
        dataCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_DATA_CELL).value)

        hasAnyValue = (Len(filePath) > 0 Or Len(sheetName) > 0 Or Len(asIsTable) > 0 Or Len(headerCell) > 0 Or Len(dataCell) > 0)
        If Not hasAnyValue Then
            GoTo NextTargetRow
        End If

        If Len(filePath) = 0 Then
            AddError errors, "target!C" & rowNum & " が空です。AsIsデータのExcelファイルパスは必須です。"
            GoTo NextTargetRow
        End If

        If Not IsAbsolutePath(filePath) Then
            AddError errors, "target!C" & rowNum & " は絶対パスで入力してください: " & filePath
            rowHasError = True
        End If

        If Len(sheetName) = 0 Then
            AddError errors, "target!D" & rowNum & " が空です。AsIsシート名は必須です。"
            rowHasError = True
        End If

        If Len(asIsTable) = 0 Then
            AddError errors, "target!E" & rowNum & " が空です。AsIsテーブル名は必須です。"
            rowHasError = True
        End If

        If Not ParseA1Address(headerCell, headerCol, headerRow) Then
            AddError errors, "target!F" & rowNum & " の開始セル(A1形式)が不正です: " & headerCell
            rowHasError = True
        End If

        If Not ParseA1Address(dataCell, dataCol, dataRow) Then
            AddError errors, "target!G" & rowNum & " の開始セル(A1形式)が不正です: " & dataCell
            rowHasError = True
        End If

        If ParseA1Address(headerCell, headerCol, headerRow) And ParseA1Address(dataCell, dataCol, dataRow) Then
            If dataRow <= headerRow Then
                AddError errors, "target 行" & rowNum & " はデータ開始セル(G列)の行番号がカラム名開始セル(F列)より下である必要があります。"
                rowHasError = True
            End If
        End If

        If rowHasError Then
            GoTo NextTargetRow
        End If

        Set targetRow = NewDictionary()
        targetRow.Add "RowNum", rowNum
        targetRow.Add "FilePath", filePath
        targetRow.Add "SheetName", sheetName
        targetRow.Add "AsIsTable", asIsTable
        targetRow.Add "HeaderCell", UCase$(headerCell)
        targetRow.Add "DataCell", UCase$(dataCell)
        targetRow.Add "OutputFilePath", BuildToBeFilePath(filePath)

        targets.Add targetRow

NextTargetRow:
    Next rowNum

    If targets.Count = 0 Then
        AddError errors, "targetシートに有効な処理対象行がありません。"
    End If
End Sub

Private Sub BindTargetsToMapping(ByRef targets As Collection, ByRef asIsToToBe As Object, ByRef errors As Collection)
    Dim item As Variant
    Dim targetRow As Object
    Dim asIsKey As String

    For Each item In targets
        Set targetRow = item
        asIsKey = NormalizeKey(CStr(targetRow("AsIsTable")))

        If Not asIsToToBe.Exists(asIsKey) Then
            AddError errors, "target 行" & CLng(targetRow("RowNum")) & " のAsIsテーブル " & CStr(targetRow("AsIsTable")) & " に対応するmapping定義がありません。"
        Else
            targetRow.Add "ToBeTableKey", CStr(asIsToToBe(asIsKey))
        End If
    Next item
End Sub

Private Sub ValidateFilesSheetsAndColumns(ByRef targets As Collection, ByRef mappingByToBe As Object, ByRef errors As Collection)
    Dim targetItem As Variant
    Dim targetRow As Object
    Dim filePath As String
    Dim outputFilePath As String
    Dim sheetName As String
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

    Set openedBooks = NewDictionary()

    On Error GoTo Cleanup

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow("FilePath"))
        outputFilePath = CStr(targetRow("OutputFilePath"))
        rowNum = CLng(targetRow("RowNum"))

        If Not FileExists(filePath) Then
            AddError errors, TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H306E) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H304C) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3A) & ChrW(&H20) & filePath
        End If

        If FileExists(outputFilePath) Then
            AddError errors, TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H306E) & ChrW(&H51FA) & ChrW(&H529B) & ChrW(&H5148) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H304C) & ChrW(&H65E2) & ChrW(&H306B) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3A) & ChrW(&H20) & outputFilePath
        End If
    Next targetItem

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow("FilePath"))
        sheetName = CStr(targetRow("SheetName"))
        rowNum = CLng(targetRow("RowNum"))

        If Not FileExists(filePath) Then
            GoTo NextTarget
        End If

        Set wb = GetOrOpenWorkbook(filePath, True, openedBooks, errors, rowNum)
        If wb Is Nothing Then
            GoTo NextTarget
        End If

        Set ws = Nothing
        If Not TryGetWorksheetFromWorkbook(wb, sheetName, ws) Then
            AddError errors, TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H306E) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H304C) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3A) & ChrW(&H20) & filePath & " / " & sheetName
            GoTo NextTarget
        End If

        toBeKey = ""
        If targetRow.Exists("ToBeTableKey") Then
            toBeKey = CStr(targetRow("ToBeTableKey"))
        End If

        If Len(toBeKey) = 0 Then
            GoTo NextTarget
        End If

        If Not mappingByToBe.Exists(toBeKey) Then
            AddError errors, TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H306B) & ChrW(&H5BFE) & ChrW(&H5FDC) & ChrW(&H3059) & ChrW(&H308B) & ChrW(&H54) & ChrW(&H6F) & ChrW(&H42) & ChrW(&H65) & ChrW(&H30C6) & ChrW(&H30FC) & ChrW(&H30D6) & ChrW(&H30EB) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
            GoTo NextTarget
        End If

        Set headerLookup = BuildHeaderLookupForValidation(ws, CStr(targetRow("HeaderCell")), errors, rowNum, filePath, sheetName)
        Set mappingRows = mappingByToBe(toBeKey)

        For Each mapItem In mappingRows
            Set mapRow = mapItem
            asIsColumn = CStr(mapRow("AsIsColumn"))

            If Len(asIsColumn) > 0 Then
                If Not headerLookup.Exists(NormalizeKey(asIsColumn)) Then
                    AddError errors, TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H3067) & ChrW(&H5FC5) & ChrW(&H8981) & ChrW(&H306A) & ChrW(&H41) & ChrW(&H73) & ChrW(&H49) & ChrW(&H73) & ChrW(&H30AB) & ChrW(&H30E9) & ChrW(&H30E0) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3A) & ChrW(&H20) & asIsColumn & " (" & filePath & " / " & sheetName & ")"
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
    Dim copiedFiles As Object
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
    Dim copyKey As String

    Set openedBooks = NewDictionary()
    Set copiedFiles = NewDictionary()

    On Error GoTo ExecError

    For Each targetItem In targets
        Set targetRow = targetItem
        sourceFilePath = CStr(targetRow("FilePath"))
        outputFilePath = CStr(targetRow("OutputFilePath"))
        copyKey = NormalizeKey(sourceFilePath)

        If Not copiedFiles.Exists(copyKey) Then
            FileCopy sourceFilePath, outputFilePath
            copiedFiles.Add copyKey, outputFilePath
            WriteLog LogLevelInfo(), ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H3092) & ChrW(&H30B3) & ChrW(&H30D4) & ChrW(&H30FC) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3A) & ChrW(&H20) & outputFilePath
        End If
    Next targetItem

    For Each targetItem In targets
        Set targetRow = targetItem
        rowNum = CLng(targetRow("RowNum"))
        outputFilePath = CStr(targetRow("OutputFilePath"))
        sheetName = CStr(targetRow("SheetName"))
        toBeKey = CStr(targetRow("ToBeTableKey"))

        Set wb = GetOrOpenWorkbookForRun(outputFilePath, openedBooks)
        Set ws = wb.Worksheets(sheetName)

        Set mappingRows = mappingByToBe(toBeKey)
        RebuildToBeSheet ws, CStr(targetRow("HeaderCell")), CStr(targetRow("DataCell")), mappingRows

        WriteLog LogLevelInfo(), TargetRowLabel(rowNum) & ChrW(&H20) & ChrW(&H3092) & ChrW(&H51E6) & ChrW(&H7406) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3A) & ChrW(&H20) & outputFilePath & " / " & sheetName
    Next targetItem

    For Each key In openedBooks.Keys
        openedBooks(key).Save
    Next key

    CloseWorkbooks openedBooks, False
    Exit Sub

ExecError:
    CloseWorkbooks openedBooks, False
    Err.Raise Err.Number, "ExecuteSort", Err.Description
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

    Set headerLookup = BuildHeaderLookupStrict(ws, headerCellAddress)

    ParseA1Address dataCellAddress, dataStartCol, dataStartRow
    ParseA1Address headerCellAddress, headerStartCol, headerRow

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

        toBeColumn = CStr(mapRow("ToBeColumn"))
        asIsColumn = CStr(mapRow("AsIsColumn"))

        outputArray(1, mapIndex) = toBeColumn

        If Len(asIsColumn) > 0 And rowCount > 0 Then
            sourceCol = CLng(headerLookup(NormalizeKey(asIsColumn)))
            columnData = ws.Range(ws.Cells(dataStartRow, sourceCol), ws.Cells(dataStartRow + rowCount - 1, sourceCol)).Value2

            For dataIndex = 1 To rowCount
                outputArray(dataIndex + 1, mapIndex) = columnData(dataIndex, 1)
            Next dataIndex
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

    ws.Range(ws.Cells(headerRow, headerStartCol), ws.Cells(lastUsedRow, lastUsedCol)).ClearContents
    ws.Range(ws.Cells(headerRow, headerStartCol), ws.Cells(headerRow + rowCount, headerStartCol + mapCount - 1)).Value2 = outputArray
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

Private Function GetOrOpenWorkbook(ByVal filePath As String, ByVal readOnly As Boolean, ByRef openedBooks As Object, ByRef errors As Collection, ByVal targetRowNum As Long) As Workbook
    Dim key As String
    Dim wb As Workbook

    key = NormalizeKey(filePath)

    If openedBooks.Exists(key) Then
        Set GetOrOpenWorkbook = openedBooks(key)
        Exit Function
    End If

    On Error GoTo OpenError
    Set wb = Workbooks.Open(Filename:=filePath, readOnly:=readOnly, UpdateLinks:=0)
    openedBooks.Add key, wb
    Set GetOrOpenWorkbook = wb
    Exit Function

OpenError:
    AddError errors, "target 行" & targetRowNum & " のファイルを開けません: " & filePath & " / " & Err.Description
    Err.Clear
End Function

Private Function GetOrOpenWorkbookForRun(ByVal filePath As String, ByRef openedBooks As Object) As Workbook
    Dim key As String
    Dim wb As Workbook

    key = NormalizeKey(filePath)

    If openedBooks.Exists(key) Then
        Set GetOrOpenWorkbookForRun = openedBooks(key)
        Exit Function
    End If

    Set wb = Workbooks.Open(Filename:=filePath, readOnly:=False, UpdateLinks:=0)
    openedBooks.Add key, wb
    Set GetOrOpenWorkbookForRun = wb
End Function

Private Function TryGetThisWorkbookSheet(ByVal sheetName As String, ByRef ws As Worksheet, ByRef errors As Collection) As Boolean
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

Private Function TryGetWorksheetFromWorkbook(ByVal wb As Workbook, ByVal sheetName As String, ByRef ws As Worksheet) As Boolean
    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    TryGetWorksheetFromWorkbook = Not (ws Is Nothing)
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = TryGetWorksheetFromWorkbook(wb, sheetName, ws)
End Function

Private Sub CloseWorkbooks(ByVal openedBooks As Object, ByVal saveChanges As Boolean)
    Dim key As Variant

    On Error Resume Next
    For Each key In openedBooks.Keys
        openedBooks(key).Close saveChanges:=saveChanges
    Next key
    On Error GoTo 0
End Sub

Private Function NewDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set NewDictionary = dict
End Function

Private Sub AddError(ByRef errors As Collection, ByVal message As String)
    errors.Add message
End Sub

Private Sub ReportValidationErrors(ByRef errors As Collection)
    Dim item As Variant

    WriteLog LogLevelError(), "入力チェックでエラーが " & errors.Count & " 件見つかりました。"

    For Each item In errors
        WriteLog LogLevelError(), CStr(item)
    Next item

    MsgBox ChrW(&H5165) & ChrW(&H529B) & ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H3067) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H304C) & ChrW(&H20) & errors.Count & ChrW(&H20) & ChrW(&H4EF6) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002) & ChrW(&H6C) & ChrW(&H6F) & ChrW(&H67) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002), vbCritical + vbOKOnly
End Sub

Private Function TrimSafe(ByVal value As Variant) As String
    If IsError(value) Then
        TrimSafe = ""
    Else
        TrimSafe = Trim$(CStr(value))
    End If
End Function

Private Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = Trim$(value)
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim found As String
    found = Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
    FileExists = (Len(found) > 0)
End Function

Private Function BuildToBeFilePath(ByVal sourceFilePath As String) As String
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

Private Function BuildExecuteConfirmMessage() As String
    BuildExecuteConfirmMessage = ChrW(&H4E26) & ChrW(&H3073) & ChrW(&H66FF) & ChrW(&H3048) & ChrW(&H3092) & ChrW(&H5B9F) & ChrW(&H884C) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & "?"
End Function

Private Function LogLevelInfo() As String
    LogLevelInfo = ChrW(&H60C5) & ChrW(&H5831)
End Function

Private Function LogLevelError() As String
    LogLevelError = ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
End Function

Private Function TargetRowLabel(ByVal rowNum As Long) As String
    TargetRowLabel = ChrW(&H5BFE) & ChrW(&H8C61) & ChrW(&H884C) & CStr(rowNum)
End Function

Private Function ParseA1Address(ByVal address As String, ByRef colNumber As Long, ByRef rowNumber As Long) As Boolean
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

Private Function ColumnLettersToNumber(ByVal letters As String) As Long
    Dim index As Long
    Dim result As Long

    result = 0
    For index = 1 To Len(letters)
        result = result * 26 + (Asc(Mid$(letters, index, 1)) - Asc("A") + 1)
    Next index

    ColumnLettersToNumber = result
End Function

Private Function IsAbsolutePath(ByVal filePath As String) As Boolean
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

Private Function GetLastRowInColumns(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long) As Long
    Dim maxRow As Long
    maxRow = 1

    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col1).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col2).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col3).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col4).End(xlUp).Row)

    GetLastRowInColumns = maxRow
End Function

Private Function GetLastRowInColumns5(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long, ByVal col5 As Long) As Long
    Dim maxRow As Long
    maxRow = 1

    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col1).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col2).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col3).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col4).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, col5).End(xlUp).Row)

    GetLastRowInColumns5 = maxRow
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range

    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = found.Row
    End If
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim found As Range

    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = found.Column
    End If
End Function

Private Sub PrepareApplication()
    gPrevScreenUpdating = Application.ScreenUpdating
    gPrevEnableEvents = Application.EnableEvents
    gPrevDisplayAlerts = Application.DisplayAlerts
    gPrevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreApplication()
    On Error Resume Next
    Application.ScreenUpdating = gPrevScreenUpdating
    Application.EnableEvents = gPrevEnableEvents
    Application.DisplayAlerts = gPrevDisplayAlerts
    Application.Calculation = gPrevCalculation
    On Error GoTo 0
End Sub

Private Sub InitLog()
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

Private Sub WriteLog(ByVal level As String, ByVal message As String)
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

