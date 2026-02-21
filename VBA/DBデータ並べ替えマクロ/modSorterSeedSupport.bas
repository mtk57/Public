Attribute VB_Name = "modSorterSeedSupport"
Option Explicit
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

