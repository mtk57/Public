Attribute VB_Name = "modAsIsToToBeSorter"
Option Explicit

Private Const VER = "1.0.0"

Private Const SHEET_MAPPING As String = "mapping"
Private Const SHEET_TARGET As String = "target"
Private Const SHEET_MACRO As String = "log"
Private Const SHEET_TOOL As String = "macro"

Private Const TOBE_FILE_SUFFIX As String = "_TOBE"
Private Const START_ROW As Long = 5

' mapping sheet columns
Private Const COL_MAP_TOBE_TABLE As Long = 2   ' B
Private Const COL_MAP_TOBE_TABLE_LOGICAL As Long = 3 ' C
Private Const COL_MAP_TOBE_COLUMN As Long = 4  ' D
Private Const COL_MAP_TOBE_COLUMN_LOGICAL As Long = 5 ' E
Private Const COL_MAP_TOBE_KEY As Long = 6 ' F
Private Const COL_MAP_TOBE_DATA_TYPE As Long = 8 ' H
Private Const COL_MAP_TOBE_DIGITS As Long = 9 ' I
Private Const COL_MAP_TOBE_SCALE As Long = 10 ' J
Private Const COL_MAP_TOBE_DESC As Long = 12 ' L
Private Const COL_MAP_TOBE_NOTE As Long = 13 ' M
Private Const COL_MAP_ASIS_TABLE As Long = 15  ' O
Private Const COL_MAP_ASIS_COLUMN As Long = 17 ' Q

' target sheet columns
Private Const COL_TARGET_MAPPING_SHEET As Long = 2 ' B
Private Const COL_TARGET_FILE As Long = 3       ' C
Private Const COL_TARGET_SHEET As Long = 4      ' D
Private Const COL_TARGET_ASIS_TABLE As Long = 5 ' E
Private Const COL_TARGET_HEADER_CELL As Long = 6 ' F
Private Const COL_TARGET_DATA_CELL As Long = 7   ' G

' macro sheet input
Private Const MACRO_CFG_COL As Long = 2 ' B
Private Const MACRO_CFG_ROW_FILE_PATH As Long = 20
Private Const MACRO_CFG_ROW_TABLE_PHYS As Long = 21
Private Const MACRO_CFG_ROW_TABLE_LOGICAL As Long = 22
Private Const MACRO_CFG_ROW_COL_PHYS As Long = 23
Private Const MACRO_CFG_ROW_COL_LOGICAL As Long = 24
Private Const MACRO_CFG_ROW_KEY As Long = 25
Private Const MACRO_CFG_ROW_DATA_TYPE As Long = 26
Private Const MACRO_CFG_ROW_DIGITS As Long = 27
Private Const MACRO_CFG_ROW_SCALE As Long = 28
Private Const MACRO_CFG_ROW_DESC As Long = 29
Private Const MACRO_CFG_ROW_NOTE As Long = 30
Private Const MACRO_CFG_ROW_DELETE_FLAG As Long = 31
Private Const MACRO_CFG_ROW_SHEETS_START As Long = 32

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

    ValidateTarget targets, errors
    ValidateMappingsForTargets targets, mappingByToBe, asIsToToBe, errors
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
    MsgBox "予期しないエラーが発生しました。logシートを確認してください。", vbCritical + vbOKOnly
    RestoreApplication
End Sub

Public Sub ToBeMappingSeedSupport()
    CreateToBeMappingSeedSupport
End Sub

Public Sub CreateToBeMappingSeedSupport()
    Dim errors As Collection
    Dim settings As Object
    Dim targetSheets As Collection
    Dim srcWb As Workbook
    Dim wsOutput As Worksheet
    Dim createdCount As Long

    On Error GoTo FatalError

    PrepareApplication
    InitLog
    WriteLog LogLevelInfo(), JpSeedStartMessage()

    Set errors = New Collection
    Set settings = NewDictionary()
    Set targetSheets = New Collection

    ValidateSeedSupportInput settings, targetSheets, errors
    If errors.Count > 0 Then
        ReportValidationErrors errors
        GoTo SafeExit
    End If

    Set srcWb = Workbooks.Open(Filename:=CStr(settings("DefFilePath")), ReadOnly:=True, UpdateLinks:=0)
    ValidateSeedSourceSheets srcWb, targetSheets, errors
    If errors.Count > 0 Then
        ReportValidationErrors errors
        GoTo SafeExit
    End If

    Set wsOutput = CreateSeedOutputSheet()
    WriteSeedSupportRows srcWb, targetSheets, settings, wsOutput, createdCount, errors

    If errors.Count > 0 Then
        ReportValidationErrors errors
        GoTo SafeExit
    End If

    WriteLog LogLevelInfo(), JpSeedCompletedMessage(wsOutput.Name, createdCount)
    MsgBox JpSeedCompletedMessage(wsOutput.Name, createdCount), vbInformation + vbOKOnly

SafeExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then
        srcWb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    RestoreApplication
    Exit Sub

FatalError:
    WriteLog LogLevelError(), JpSeedFatalPrefix() & " (" & Err.Number & ") " & Err.Description
    MsgBox JpSeedFatalDialogMessage(), vbCritical + vbOKOnly

    On Error Resume Next
    If Not srcWb Is Nothing Then
        srcWb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    RestoreApplication
End Sub

Private Sub ValidateSeedSupportInput(ByRef settings As Object, ByRef targetSheets As Collection, ByRef errors As Collection)
    Dim wsTool As Worksheet
    Dim defFilePath As String
    Dim lastRow As Long
    Dim rowNum As Long
    Dim sheetName As String
    Dim namesSeen As Object
    Dim key As String

    If Not TryGetThisWorkbookSheet(SHEET_TOOL, wsTool, errors) Then
        Exit Sub
    End If

    defFilePath = TrimSafe(wsTool.Cells(MACRO_CFG_ROW_FILE_PATH, MACRO_CFG_COL).Value)
    If Len(defFilePath) = 0 Then
        AddError errors, BuildMacroCellLabel(MACRO_CFG_ROW_FILE_PATH) & JpRequiredSuffix()
    Else
        If Not IsAbsolutePath(defFilePath) Then
            AddError errors, BuildMacroCellLabel(MACRO_CFG_ROW_FILE_PATH) & JpAbsolutePathSuffix() & defFilePath
        End If

        If Not FileExists(defFilePath) Then
            AddError errors, BuildMacroCellLabel(MACRO_CFG_ROW_FILE_PATH) & JpDefFileNotFoundSuffix() & defFilePath
        End If
    End If
    SetSettingValue settings, "DefFilePath", defFilePath

    ReadAddressSetting wsTool, MACRO_CFG_ROW_TABLE_PHYS, "TablePhys", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_TABLE_LOGICAL, "TableLogical", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_COL_PHYS, "ColPhys", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_COL_LOGICAL, "ColLogical", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_KEY, "Key", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_DATA_TYPE, "DataType", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_DIGITS, "Digits", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_SCALE, "Scale", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_DESC, "Desc", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_NOTE, "Note", True, settings, errors
    ReadAddressSetting wsTool, MACRO_CFG_ROW_DELETE_FLAG, "DeleteFlag", True, settings, errors

    Set namesSeen = NewDictionary()
    lastRow = wsTool.Cells(wsTool.Rows.Count, MACRO_CFG_COL).End(xlUp).Row

    If lastRow >= MACRO_CFG_ROW_SHEETS_START Then
        For rowNum = MACRO_CFG_ROW_SHEETS_START To lastRow
            sheetName = TrimSafe(wsTool.Cells(rowNum, MACRO_CFG_COL).Value)
            If Len(sheetName) > 0 Then
                key = NormalizeKey(sheetName)
                If namesSeen.Exists(key) Then
                    AddError errors, BuildMacroCellLabel(rowNum) & JpDuplicateSheetNameSuffix() & sheetName
                Else
                    namesSeen.Add key, True
                    targetSheets.Add sheetName
                End If
            End If
        Next rowNum
    End If

    If targetSheets.Count = 0 Then
        AddError errors, JpNoSheetNamesMessage()
    End If
End Sub

Private Sub ReadAddressSetting(ByVal wsTool As Worksheet, ByVal configRow As Long, ByVal keyPrefix As String, ByVal required As Boolean, ByRef settings As Object, ByRef errors As Collection)
    Dim cellValue As String
    Dim colNumber As Long
    Dim rowNumber As Long

    cellValue = TrimSafe(wsTool.Cells(configRow, MACRO_CFG_COL).Value)

    If Len(cellValue) = 0 Then
        If required Then
            AddError errors, BuildMacroCellLabel(configRow) & JpRequiredSuffix()
        End If

        SetSettingValue settings, keyPrefix & "Col", 0
        SetSettingValue settings, keyPrefix & "Row", 0
        Exit Sub
    End If

    If Not ParseA1Address(cellValue, colNumber, rowNumber) Then
        AddError errors, BuildMacroCellLabel(configRow) & JpInvalidA1Suffix() & cellValue

        SetSettingValue settings, keyPrefix & "Col", 0
        SetSettingValue settings, keyPrefix & "Row", 0
        Exit Sub
    End If

    SetSettingValue settings, keyPrefix & "Col", colNumber
    SetSettingValue settings, keyPrefix & "Row", rowNumber
End Sub

Private Sub ValidateSeedSourceSheets(ByVal srcWb As Workbook, ByVal targetSheets As Collection, ByRef errors As Collection)
    Dim sheetItem As Variant
    Dim ws As Worksheet

    For Each sheetItem In targetSheets
        Set ws = Nothing
        If Not TryGetWorksheetFromWorkbook(srcWb, CStr(sheetItem), ws) Then
            AddError errors, CStr(sheetItem) & JpDefSheetNotFoundSuffix() & srcWb.FullName
        End If
    Next sheetItem
End Sub

Private Function CreateSeedOutputSheet() As Worksheet
    Dim wsMapping As Worksheet
    Dim wsOutput As Worksheet
    Dim sheetName As String

    Set wsMapping = ThisWorkbook.Worksheets(SHEET_MAPPING)

    Set wsOutput = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sheetName = BuildUniqueMappingSeedSheetName()
    wsOutput.Name = sheetName

    wsMapping.Rows("1:" & CStr(START_ROW - 1)).Copy Destination:=wsOutput.Rows(1)

    Set CreateSeedOutputSheet = wsOutput
End Function

Private Function BuildUniqueMappingSeedSheetName() As String
    Dim baseName As String
    Dim candidate As String
    Dim index As Long

    baseName = "mapping_seed_" & Format$(Now, "yyyymmdd_hhnnss")

    For index = 0 To 99
        If index = 0 Then
            candidate = Left$(baseName, 31)
        Else
            candidate = Left$(baseName, 28) & "_" & Format$(index, "00")
        End If

        If Not WorksheetExists(ThisWorkbook, candidate) Then
            BuildUniqueMappingSeedSheetName = candidate
            Exit Function
        End If
    Next index

    Err.Raise vbObjectError + 2101, "BuildUniqueMappingSeedSheetName", "出力シート名を作成できません。"
End Function

Private Sub WriteSeedSupportRows(ByVal srcWb As Workbook, ByVal targetSheets As Collection, ByVal settings As Object, ByVal wsOutput As Worksheet, ByRef createdCount As Long, ByRef errors As Collection)
    Dim sheetItem As Variant
    Dim wsSrc As Worksheet
    Dim outRow As Long
    Dim rowOffset As Long
    Dim copiedInSheet As Long
    Dim toBeCol As String
    Dim deleteFlag As String
    Dim tablePhys As String

    outRow = START_ROW

    For Each sheetItem In targetSheets
        Set wsSrc = srcWb.Worksheets(CStr(sheetItem))

        tablePhys = TrimSafe(GetFixedCellValue(wsSrc, settings, "TablePhys"))
        If Len(tablePhys) = 0 Then
            AddError errors, CStr(sheetItem) & JpTablePhysEmptySuffix() & srcWb.FullName
            GoTo NextSheet
        End If

        copiedInSheet = 0
        rowOffset = 0

        Do
            toBeCol = TrimSafe(GetRowCellValue(wsSrc, settings, "ColPhys", rowOffset))
            If Len(toBeCol) = 0 Then
                Exit Do
            End If

            deleteFlag = UCase$(TrimSafe(GetRowCellValue(wsSrc, settings, "DeleteFlag", rowOffset)))
            If deleteFlag <> "X" Then
                wsOutput.Cells(outRow, COL_MAP_TOBE_TABLE).Value = GetFixedCellValue(wsSrc, settings, "TablePhys")
                wsOutput.Cells(outRow, COL_MAP_TOBE_TABLE_LOGICAL).Value = GetFixedCellValue(wsSrc, settings, "TableLogical")
                wsOutput.Cells(outRow, COL_MAP_TOBE_COLUMN).Value = GetRowCellValue(wsSrc, settings, "ColPhys", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_COLUMN_LOGICAL).Value = GetRowCellValue(wsSrc, settings, "ColLogical", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_KEY).Value = GetRowCellValue(wsSrc, settings, "Key", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_DATA_TYPE).Value = GetRowCellValue(wsSrc, settings, "DataType", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_DIGITS).Value = GetRowCellValue(wsSrc, settings, "Digits", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_SCALE).Value = GetRowCellValue(wsSrc, settings, "Scale", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_DESC).Value = GetRowCellValue(wsSrc, settings, "Desc", rowOffset)
                wsOutput.Cells(outRow, COL_MAP_TOBE_NOTE).Value = GetRowCellValue(wsSrc, settings, "Note", rowOffset)

                outRow = outRow + 1
                copiedInSheet = copiedInSheet + 1
            End If

            rowOffset = rowOffset + 1
        Loop

        WriteLog LogLevelInfo(), JpSheetProcessedPrefix() & CStr(sheetItem) & " (" & copiedInSheet & ")"

NextSheet:
    Next sheetItem

    createdCount = outRow - START_ROW
End Sub

Private Function GetFixedCellValue(ByVal ws As Worksheet, ByVal settings As Object, ByVal keyPrefix As String) As Variant
    Dim rowNumber As Long
    Dim colNumber As Long

    rowNumber = CLng(settings(keyPrefix & "Row"))
    colNumber = CLng(settings(keyPrefix & "Col"))

    If rowNumber <= 0 Or colNumber <= 0 Then
        GetFixedCellValue = vbNullString
        Exit Function
    End If

    GetFixedCellValue = ws.Cells(rowNumber, colNumber).Value
End Function

Private Function GetRowCellValue(ByVal ws As Worksheet, ByVal settings As Object, ByVal keyPrefix As String, ByVal rowOffset As Long) As Variant
    Dim rowNumber As Long
    Dim colNumber As Long

    rowNumber = CLng(settings(keyPrefix & "Row"))
    colNumber = CLng(settings(keyPrefix & "Col"))

    If rowNumber <= 0 Or colNumber <= 0 Then
        GetRowCellValue = vbNullString
        Exit Function
    End If

    GetRowCellValue = ws.Cells(rowNumber + rowOffset, colNumber).Value
End Function

Private Sub SetSettingValue(ByRef settings As Object, ByVal key As String, ByVal value As Variant)
    If settings.Exists(key) Then
        settings(key) = value
    Else
        settings.Add key, value
    End If
End Sub

Private Function BuildMacroCellLabel(ByVal rowNum As Long) As String
    BuildMacroCellLabel = "macro!B" & CStr(rowNum)
End Function

Private Function JpRequiredSuffix() As String
    JpRequiredSuffix = " は必須です。"
End Function

Private Function JpAbsolutePathSuffix() As String
    JpAbsolutePathSuffix = " は絶対パスで入力してください: "
End Function

Private Function JpInvalidA1Suffix() As String
    JpInvalidA1Suffix = " のセル位置(A1形式)が不正です: "
End Function

Private Function JpNoSheetNamesMessage() As String
    JpNoSheetNamesMessage = "macro!B32以降にシート名を1件以上入力してください。"
End Function

Private Function JpDefSheetNotFoundSuffix() As String
    JpDefSheetNotFoundSuffix = " ToBeテーブル定義ファイルにシートが存在しません: "
End Function

Private Function JpDefFileNotFoundSuffix() As String
    JpDefFileNotFoundSuffix = " ToBeテーブル定義ファイルが存在しません: "
End Function

Private Function JpDuplicateSheetNameSuffix() As String
    JpDuplicateSheetNameSuffix = " シート名が重複しています: "
End Function

Private Function JpTablePhysEmptySuffix() As String
    JpTablePhysEmptySuffix = " テーブル物理名が空のため処理できません: "
End Function

Private Function JpSeedStartMessage() As String
    JpSeedStartMessage = "ToBeマッピング元ネタ作成支援を開始します。"
End Function

Private Function JpSeedCompletedMessage(ByVal outputSheetName As String, ByVal createdCount As Long) As String
    JpSeedCompletedMessage = "ToBeマッピング元ネタ作成支援が完了しました。" & vbCrLf & JpSeedOutputSheetPrefix() & outputSheetName & vbCrLf & JpSeedCreatedCountPrefix() & CStr(createdCount)
End Function

Private Function JpSeedOutputSheetPrefix() As String
    JpSeedOutputSheetPrefix = "出力シート名: "
End Function

Private Function JpSeedCreatedCountPrefix() As String
    JpSeedCreatedCountPrefix = "作成件数: "
End Function

Private Function JpSheetProcessedPrefix() As String
    JpSheetProcessedPrefix = "シート処理完了: "
End Function

Private Function JpSeedFatalPrefix() As String
    JpSeedFatalPrefix = "ToBeマッピング元ネタ作成支援でエラー"
End Function

Private Function JpSeedFatalDialogMessage() As String
    JpSeedFatalDialogMessage = "予期しないエラーが発生しました。logシートを確認してください。"
End Function

Private Function JpNoDataRowsSuffix() As String
    JpNoDataRowsSuffix = "にデータ行がありません。"
End Function

Private Function JpNoValidRowsSuffix() As String
    JpNoValidRowsSuffix = "に有効な行がありません。"
End Function

Private Function JpMappingAsIsPairError(ByVal mappingSheetName As String, ByVal rowNum As Long) As String
    JpMappingAsIsPairError = "[" & mappingSheetName & "] 行" & rowNum & "のAsIsテーブル(O列)とAsIsカラム(Q列)は両方入力または両方空欄にしてください。"
End Function

Private Function JpDuplicateToBeColumnError(ByVal mappingSheetName As String, ByVal rowNum As Long, ByVal toBeTable As String, ByVal toBeColumn As String) As String
    JpDuplicateToBeColumnError = "[" & mappingSheetName & "] 行" & rowNum & "でToBeカラムが重複しています: " & toBeTable & "." & toBeColumn
End Function

Private Function JpAsIsToToBeConflictError(ByVal mappingSheetName As String, ByVal rowNum As Long, ByVal asIsTable As String) As String
    JpAsIsToToBeConflictError = "[" & mappingSheetName & "] 行" & rowNum & "でAsIsテーブルの対応先が重複しています: " & asIsTable
End Function

Private Function JpToBeToAsIsConflictError(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    JpToBeToAsIsConflictError = "[" & mappingSheetName & "] ToBeテーブルに複数のAsIsテーブルが定義されています: " & toBeTable
End Function

Private Function JpToBeNoAsIsError(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    JpToBeNoAsIsError = "[" & mappingSheetName & "] ToBeテーブルにAsIsテーブルが1件も定義されていません: " & toBeTable
End Function

Private Function JpTargetDataStartRowOrderSuffix() As String
    JpTargetDataStartRowOrderSuffix = "ではデータ開始セル(G列)はカラム名開始セル(F列)より下の行を指定してください。"
End Function

Private Function JpTargetAsIsMappingNotFoundError(ByVal rowNum As Long, ByVal asIsTable As String, ByVal mappingSheetName As String) As String
    JpTargetAsIsMappingNotFoundError = TargetRowLabel(rowNum) & "でAsIsテーブルに対応するマッピングが見つかりません: " & asIsTable & " (マッピングシート: " & mappingSheetName & ")"
End Function

Private Function JpTargetToBeMappingNotFoundError(ByVal rowNum As Long, ByVal mappingSheetName As String) As String
    JpTargetToBeMappingNotFoundError = TargetRowLabel(rowNum) & "でToBeテーブル定義が見つかりません (マッピングシート: " & mappingSheetName & ")。"
End Function

Private Sub ValidateMappingsForTargets(ByRef targets As Collection, ByRef mappingByToBe As Object, ByRef asIsToToBe As Object, ByRef errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim mappingSheetName As String
    Dim filePath As String
    Dim sheetName As String
    Dim asIsTable As String
    Dim headerCell As String
    Dim dataCell As String
    Dim hasAnyValue As Boolean
    Dim sheetKey As String
    Dim validatedSheets As Object

    If Not TryGetThisWorkbookSheet(SHEET_TARGET, ws, errors) Then
        Exit Sub
    End If

    Set validatedSheets = NewDictionary()
    lastRow = GetLastRowInColumns6(ws, COL_TARGET_MAPPING_SHEET, COL_TARGET_FILE, COL_TARGET_SHEET, COL_TARGET_ASIS_TABLE, COL_TARGET_HEADER_CELL, COL_TARGET_DATA_CELL)

    If lastRow < START_ROW Then
        Exit Sub
    End If

    For rowNum = START_ROW To lastRow
        mappingSheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_MAPPING_SHEET).Value)
        filePath = TrimSafe(ws.Cells(rowNum, COL_TARGET_FILE).Value)
        sheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_SHEET).Value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_TARGET_ASIS_TABLE).Value)
        headerCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_HEADER_CELL).Value)
        dataCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_DATA_CELL).Value)

        hasAnyValue = (Len(mappingSheetName) > 0 Or Len(filePath) > 0 Or Len(sheetName) > 0 Or Len(asIsTable) > 0 Or Len(headerCell) > 0 Or Len(dataCell) > 0)
        If Not hasAnyValue Then
            GoTo NextTargetRow
        End If

        If Len(mappingSheetName) = 0 Then
            GoTo NextTargetRow
        End If

        sheetKey = NormalizeKey(mappingSheetName)
        If Not validatedSheets.Exists(sheetKey) Then
            validatedSheets.Add sheetKey, mappingSheetName
            ValidateMappingSheet mappingSheetName, mappingByToBe, asIsToToBe, errors
        End If

NextTargetRow:
    Next rowNum
End Sub


Private Sub ValidateMappingSheet(ByVal mappingSheetName As String, ByRef mappingByToBe As Object, ByRef asIsToToBe As Object, ByRef errors As Collection)
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

    For rowNum = START_ROW To lastRow
        toBeTable = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_TABLE).Value)
        toBeColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_TOBE_COLUMN).Value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_TABLE).Value)
        asIsColumn = TrimSafe(ws.Cells(rowNum, COL_MAP_ASIS_COLUMN).Value)

        hasAnyValue = (Len(toBeTable) > 0 Or Len(toBeColumn) > 0 Or Len(asIsTable) > 0 Or Len(asIsColumn) > 0)
        If Not hasAnyValue Then
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
        mapRow.Add "RowNum", rowNum
        mapRow.Add "ToBeTable", toBeTable
        mapRow.Add "ToBeColumn", toBeColumn
        mapRow.Add "AsIsTable", asIsTable
        mapRow.Add "AsIsColumn", asIsColumn

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

    If validRowCount = 0 Then
        AddError errors, "[" & mappingSheetName & "]" & JpNoValidRowsSuffix()
        Exit Sub
    End If

    For Each key In sheetToBeKeys.Keys
        If Not toBeToAsIsTable.Exists(CStr(key)) Then
            Set mapRows = mappingByToBe(CStr(key))
            Set mapRow = mapRows(1)
            AddError errors, JpToBeNoAsIsError(mappingSheetName, CStr(mapRow("ToBeTable")))
        End If
    Next key
End Sub

Private Sub ValidateTarget(ByRef targets As Collection, ByRef errors As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim mappingSheetName As String
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

    lastRow = GetLastRowInColumns6(ws, COL_TARGET_MAPPING_SHEET, COL_TARGET_FILE, COL_TARGET_SHEET, COL_TARGET_ASIS_TABLE, COL_TARGET_HEADER_CELL, COL_TARGET_DATA_CELL)
    If lastRow < START_ROW Then
        AddError errors, "target" & JpNoDataRowsSuffix()
        Exit Sub
    End If

    For rowNum = START_ROW To lastRow
        rowHasError = False
        mappingSheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_MAPPING_SHEET).Value)
        filePath = TrimSafe(ws.Cells(rowNum, COL_TARGET_FILE).Value)
        sheetName = TrimSafe(ws.Cells(rowNum, COL_TARGET_SHEET).Value)
        asIsTable = TrimSafe(ws.Cells(rowNum, COL_TARGET_ASIS_TABLE).Value)
        headerCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_HEADER_CELL).Value)
        dataCell = TrimSafe(ws.Cells(rowNum, COL_TARGET_DATA_CELL).Value)

        hasAnyValue = (Len(mappingSheetName) > 0 Or Len(filePath) > 0 Or Len(sheetName) > 0 Or Len(asIsTable) > 0 Or Len(headerCell) > 0 Or Len(dataCell) > 0)
        If Not hasAnyValue Then
            GoTo NextTargetRow
        End If

        If Len(mappingSheetName) = 0 Then
            AddError errors, "target!B" & rowNum & JpRequiredSuffix()
            rowHasError = True
        End If

        If Len(filePath) = 0 Then
            AddError errors, "target!C" & rowNum & JpRequiredSuffix()
            GoTo NextTargetRow
        End If

        If Not IsAbsolutePath(filePath) Then
            AddError errors, "target!C" & rowNum & JpAbsolutePathSuffix() & filePath
            rowHasError = True
        End If

        If Len(sheetName) = 0 Then
            AddError errors, "target!D" & rowNum & JpRequiredSuffix()
            rowHasError = True
        End If

        If Len(asIsTable) = 0 Then
            AddError errors, "target!E" & rowNum & JpRequiredSuffix()
            rowHasError = True
        End If

        If Not ParseA1Address(headerCell, headerCol, headerRow) Then
            AddError errors, "target!F" & rowNum & JpInvalidA1Suffix() & headerCell
            rowHasError = True
        End If

        If Not ParseA1Address(dataCell, dataCol, dataRow) Then
            AddError errors, "target!G" & rowNum & JpInvalidA1Suffix() & dataCell
            rowHasError = True
        End If

        If ParseA1Address(headerCell, headerCol, headerRow) And ParseA1Address(dataCell, dataCol, dataRow) Then
            If dataRow <= headerRow Then
                AddError errors, TargetRowLabel(rowNum) & JpTargetDataStartRowOrderSuffix()
                rowHasError = True
            End If
        End If

        If rowHasError Then
            GoTo NextTargetRow
        End If

        Set targetRow = NewDictionary()
        targetRow.Add "RowNum", rowNum
        targetRow.Add "MappingSheetName", mappingSheetName
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
        AddError errors, "target" & JpNoValidRowsSuffix()
    End If
End Sub

Private Sub BindTargetsToMapping(ByRef targets As Collection, ByRef asIsToToBe As Object, ByRef errors As Collection)
    Dim item As Variant
    Dim targetRow As Object
    Dim asIsKey As String
    Dim mappingSheetName As String

    For Each item In targets
        Set targetRow = item
        mappingSheetName = CStr(targetRow("MappingSheetName"))
        asIsKey = BuildAsIsMappingKey(mappingSheetName, CStr(targetRow("AsIsTable")))

        If Not asIsToToBe.Exists(asIsKey) Then
            AddError errors, JpTargetAsIsMappingNotFoundError(CLng(targetRow("RowNum")), CStr(targetRow("AsIsTable")), mappingSheetName)
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

    Set openedBooks = NewDictionary()

    On Error GoTo Cleanup

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow("FilePath"))
        outputFilePath = CStr(targetRow("OutputFilePath"))
        rowNum = CLng(targetRow("RowNum"))

        If Not FileExists(filePath) Then
            AddError errors, TargetRowLabel(rowNum) & " のファイルが存在しません: " & filePath
        End If

        If FileExists(outputFilePath) Then
            AddError errors, TargetRowLabel(rowNum) & " の出力先ファイルが既に存在します: " & outputFilePath
        End If
    Next targetItem

    For Each targetItem In targets
        Set targetRow = targetItem
        filePath = CStr(targetRow("FilePath"))
        sheetName = CStr(targetRow("SheetName"))
        mappingSheetName = CStr(targetRow("MappingSheetName"))
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
            AddError errors, TargetRowLabel(rowNum) & " のシートが存在しません: " & filePath & " / " & sheetName
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
            AddError errors, JpTargetToBeMappingNotFoundError(rowNum, mappingSheetName)
            GoTo NextTarget
        End If

        Set headerLookup = BuildHeaderLookupForValidation(ws, CStr(targetRow("HeaderCell")), errors, rowNum, filePath, sheetName)
        Set mappingRows = mappingByToBe(toBeKey)

        For Each mapItem In mappingRows
            Set mapRow = mapItem
            asIsColumn = CStr(mapRow("AsIsColumn"))

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
            WriteLog LogLevelInfo(), "ファイルをコピーしました: " & outputFilePath
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

        WriteLog LogLevelInfo(), TargetRowLabel(rowNum) & " を処理しました: " & outputFilePath & " / " & sheetName
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

    MsgBox "入力チェックでエラーが " & errors.Count & " 件見つかりました。logシートを確認してください。", vbCritical + vbOKOnly
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

Private Function BuildToBeMappingKey(ByVal mappingSheetName As String, ByVal toBeTable As String) As String
    BuildToBeMappingKey = NormalizeKey(mappingSheetName) & "|" & NormalizeKey(toBeTable)
End Function

Private Function BuildAsIsMappingKey(ByVal mappingSheetName As String, ByVal asIsTable As String) As String
    BuildAsIsMappingKey = NormalizeKey(mappingSheetName) & "|" & NormalizeKey(asIsTable)
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
    BuildExecuteConfirmMessage = "並び替えを実行しますか?"
End Function

Private Function LogLevelInfo() As String
    LogLevelInfo = "情報"
End Function

Private Function LogLevelError() As String
    LogLevelError = "エラー"
End Function

Private Function TargetRowLabel(ByVal rowNum As Long) As String
    TargetRowLabel = "対象行" & CStr(rowNum)
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

Private Function GetLastRowInColumns6(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal col3 As Long, ByVal col4 As Long, ByVal col5 As Long, ByVal col6 As Long) As Long
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

