Attribute VB_Name = "main"
Option Explicit

Private Const VER = "1.0.7"

Public Sub RunRenumbering()
    On Error GoTo ErrHandler
    Dim configSheet As Worksheet
    Set configSheet = ActiveSheet
    
    Dim fileInput As String
    Dim sheetName As String
    Dim majorColumnLetter As String
    Dim middleColumnLetter As String
    Dim minorColumnLetter As String
    Dim startNumber As Long
    Dim startNumberValue As Variant
    Dim startNumberDouble As Double
    
    fileInput = Trim$(CStr(configSheet.Range("C6").value))
    sheetName = Trim$(CStr(configSheet.Range("C7").value))
    majorColumnLetter = UCase$(Trim$(CStr(configSheet.Range("C8").value)))
    middleColumnLetter = UCase$(Trim$(CStr(configSheet.Range("C9").value)))
    minorColumnLetter = UCase$(Trim$(CStr(configSheet.Range("C10").value)))
    
    If Len(fileInput) = 0 Then
        MsgBox "C6セルに処理対象のファイル名（またはパス）を入力してください。", vbExclamation
        Exit Sub
    End If
    
    If Len(sheetName) = 0 Then
        MsgBox "C7セルに処理対象のシート名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    If Len(majorColumnLetter) = 0 Then
        MsgBox "C8セルに大項番列の列記号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Dim majorColumn As Long
    majorColumn = ColumnLetterToNumber(majorColumnLetter)
    If majorColumn = 0 Then
        MsgBox "大項番列名にはA～XFDの列記号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Dim hasMiddle As Boolean
    Dim middleColumn As Long
    hasMiddle = Len(middleColumnLetter) > 0
    If hasMiddle Then
        middleColumn = ColumnLetterToNumber(middleColumnLetter)
        If middleColumn = 0 Then
            MsgBox "中項番列名にはA～XFDの列記号を入力してください。", vbExclamation
            Exit Sub
        End If
    End If
    
    Dim hasMinor As Boolean
    Dim minorColumn As Long
    hasMinor = Len(minorColumnLetter) > 0
    If hasMinor Then
        minorColumn = ColumnLetterToNumber(minorColumnLetter)
        If minorColumn = 0 Then
            MsgBox "小項番列名にはA～XFDの列記号を入力してください。", vbExclamation
            Exit Sub
        End If
    End If
    
    startNumberValue = configSheet.Range("C11").value
    If Len(Trim$(CStr(startNumberValue))) = 0 Then
        startNumber = 1
    ElseIf Not IsNumeric(startNumberValue) Then
        MsgBox "C11セルには0以上の整数を入力してください。", vbExclamation
        Exit Sub
    Else
        startNumberDouble = CDbl(startNumberValue)
        If startNumberDouble < 0 Then
            MsgBox "C11セルには0以上の整数を入力してください。", vbExclamation
            Exit Sub
        End If
        If startNumberDouble <> Fix(startNumberDouble) Then
            MsgBox "C11セルには0以上の整数を入力してください。", vbExclamation
            Exit Sub
        End If
        startNumber = CLng(startNumberDouble)
    End If
    
    Dim targetWorkbook As Workbook
    Set targetWorkbook = LocateWorkbook(fileInput)
    If targetWorkbook Is Nothing Then
        MsgBox "指定されたブックを開けませんでした。パスとファイル名を確認してください。", vbCritical
        Exit Sub
    End If
    
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(sheetName)
    On Error GoTo ErrHandler
    If targetSheet Is Nothing Then
        MsgBox "指定されたシートが存在しません。入力値を確認してください。", vbCritical
        Exit Sub
    End If
    
    targetSheet.Activate
    
    ApplyRenumbering targetSheet, majorColumn, hasMiddle, middleColumn, hasMinor, minorColumn, startNumber
    
    MsgBox "項番の振り直しが完了しました。", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Number & " : " & Err.Description, vbCritical
End Sub

Private Function LocateWorkbook(ByVal inputValue As String) As Workbook
    Dim normalizedInput As String
    normalizedInput = Trim$(inputValue)
    Dim fileNameOnly As String
    fileNameOnly = ExtractFileName(normalizedInput)
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, normalizedInput, vbTextCompare) = 0 Then
            Set LocateWorkbook = wb
            Exit Function
        End If
        
        If StrComp(wb.Name, normalizedInput, vbTextCompare) = 0 Then
            Set LocateWorkbook = wb
            Exit Function
        End If
        
        If Len(fileNameOnly) > 0 Then
            If StrComp(wb.Name, fileNameOnly, vbTextCompare) = 0 Then
                Set LocateWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
    
    Dim candidatePath As String
    candidatePath = normalizedInput
    If Len(Dir$(candidatePath)) = 0 Then
        If Len(fileNameOnly) > 0 Then
            candidatePath = ThisWorkbook.Path & Application.PathSeparator & fileNameOnly
        End If
    End If
    
    On Error Resume Next
    Set LocateWorkbook = Application.Workbooks.Open(candidatePath)
    On Error GoTo 0
End Function

Private Function ExtractFileName(ByVal pathValue As String) As String
    Dim result As String
    Dim separatorPos As Long
    Dim cleanedPath As String
    
    cleanedPath = Replace(pathValue, "/", Application.PathSeparator)
    separatorPos = InStrRev(cleanedPath, Application.PathSeparator)
    If separatorPos > 0 Then
        result = Mid$(cleanedPath, separatorPos + 1)
    Else
        result = cleanedPath
    End If
    
    ExtractFileName = result
End Function

Private Sub ApplyRenumbering(ByVal targetSheet As Worksheet, ByVal majorColumn As Long, ByVal hasMiddle As Boolean, ByVal middleColumn As Long, ByVal hasMinor As Boolean, ByVal minorColumn As Long, ByVal startNumber As Long)
    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation
    
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    lastRow = GetLastUsedRow(targetSheet, majorColumn)
    If hasMiddle Then
        lastRow = Application.WorksheetFunction.Max(lastRow, GetLastUsedRow(targetSheet, middleColumn))
    End If
    If hasMinor Then
        lastRow = Application.WorksheetFunction.Max(lastRow, GetLastUsedRow(targetSheet, minorColumn))
    End If
    
    If lastRow = 0 Then
        Application.Calculation = originalCalculation
        Application.EnableEvents = originalEnableEvents
        Application.ScreenUpdating = originalScreenUpdating
        Exit Sub
    End If
    
    Dim rowIndex As Long
    Dim currentMajor As Long
    Dim expectedMajor As Long
    Dim currentMiddle As Long
    Dim expectedMiddle As Long
    Dim expectedMinor As Long
    
    expectedMajor = startNumber
    expectedMiddle = 1
    expectedMinor = 1
    
    For rowIndex = 1 To lastRow
        If ShouldRenumberCell(targetSheet, rowIndex, majorColumn, False) Then
            WriteNormalizedNumber targetSheet.Cells(rowIndex, majorColumn), CStr(expectedMajor), False
            currentMajor = expectedMajor
            expectedMajor = expectedMajor + 1
            expectedMiddle = 1
            currentMiddle = 0
            expectedMinor = 1
        End If
        
        If hasMiddle Then
            If ShouldRenumberCell(targetSheet, rowIndex, middleColumn, True) Then
                If currentMajor = 0 Then
                    currentMajor = startNumber
                    If expectedMajor <= currentMajor Then
                        expectedMajor = currentMajor + 1
                    End If
                End If
                
                WriteNormalizedNumber targetSheet.Cells(rowIndex, middleColumn), _
                    CStr(currentMajor) & "." & CStr(expectedMiddle), True
                currentMiddle = expectedMiddle
                expectedMiddle = expectedMiddle + 1
                expectedMinor = 1
            End If
        End If
        
        If hasMinor Then
            If ShouldRenumberCell(targetSheet, rowIndex, minorColumn, True) Then
                If currentMajor = 0 Then
                    currentMajor = startNumber
                    If expectedMajor <= currentMajor Then
                        expectedMajor = currentMajor + 1
                    End If
                End If
                
                If currentMiddle = 0 Then
                    currentMiddle = 1
                    If expectedMiddle <= currentMiddle Then
                        expectedMiddle = currentMiddle + 1
                    End If
                End If
                
                WriteNormalizedNumber targetSheet.Cells(rowIndex, minorColumn), _
                    CStr(currentMajor) & "." & CStr(currentMiddle) & "." & CStr(expectedMinor), True
                expectedMinor = expectedMinor + 1
            End If
        End If
    Next rowIndex
    
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    Exit Sub
    
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    Dim errSource As String
    
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Sub

Private Function ColumnLetterToNumber(ByVal columnLetter As String) As Long
    Dim i As Long
    Dim result As Long
    Dim charCode As Long
    
    For i = 1 To Len(columnLetter)
        charCode = Asc(Mid$(columnLetter, i, 1))
        If charCode < 65 Or charCode > 90 Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        result = result * 26 + (charCode - 64)
    Next i
    
    ColumnLetterToNumber = result
End Function

Private Function GetLastUsedRow(ByVal targetSheet As Worksheet, ByVal columnIndex As Long) As Long
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, columnIndex).End(xlUp).Row
    
    If lastRow = 1 Then
        If HasCellValue(targetSheet.Cells(1, columnIndex)) Then
            GetLastUsedRow = 1
        Else
            GetLastUsedRow = 0
        End If
    Else
        GetLastUsedRow = lastRow
    End If
End Function

Private Function ShouldRenumberCell(ByVal targetSheet As Worksheet, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal allowDots As Boolean) As Boolean
    Dim targetCell As Range
    Set targetCell = targetSheet.Cells(rowIndex, columnIndex)
    
    If Not HasCellValue(targetCell) Then
        ShouldRenumberCell = False
        Exit Function
    End If
    
    Dim targetText As String
    targetText = GetCellTextForCheck(targetCell)
    If Not HasNumericPrefix(targetText, allowDots) Then
        ShouldRenumberCell = False
        Exit Function
    End If
    
    If columnIndex < targetSheet.Columns.Count Then
        Dim rightCell As Range
        Set rightCell = targetSheet.Cells(rowIndex, columnIndex + 1)
        Dim rightText As String
        rightText = GetCellTextForCheck(rightCell)
        If ShouldSkipDueToRightCell(rightText, True) Then
            ShouldRenumberCell = False
            Exit Function
        End If
    End If
    
    ShouldRenumberCell = True
End Function

Private Function GetCellTextForCheck(ByVal targetCell As Range) As String
    Dim cellFormula As String
    Dim cellText As String
    Dim rawValue As Variant
    
    On Error Resume Next
    cellFormula = targetCell.Formula
    On Error GoTo 0
    
    If Len(cellFormula) > 0 Then
        If Left$(cellFormula, 1) = "'" Then
            cellFormula = Mid$(cellFormula, 2)
            GetCellTextForCheck = cellFormula
            Exit Function
        ElseIf Left$(cellFormula, 1) <> "=" Then
            GetCellTextForCheck = cellFormula
            Exit Function
        End If
    End If
    
    cellText = targetCell.Text
    If Len(cellText) > 0 Then
        GetCellTextForCheck = cellText
        Exit Function
    End If
    
    rawValue = targetCell.value
    If IsError(rawValue) Then
        GetCellTextForCheck = ""
    Else
        GetCellTextForCheck = CStr(rawValue)
    End If
End Function

Private Function ShouldSkipDueToRightCell(ByVal cellText As String, ByVal allowDots As Boolean) As Boolean
    Dim normalized As String
    Dim remainder As String
    
    If Not ExtractNumericPrefix(cellText, allowDots, normalized, remainder) Then
        ShouldSkipDueToRightCell = False
        Exit Function
    End If
    
    If Len(remainder) = 0 Then
        ShouldSkipDueToRightCell = True
        Exit Function
    End If
    
    Dim firstChar As String
    firstChar = Left$(remainder, 1)
    If firstChar = " " Or firstChar = ChrW$(&H3000) Or firstChar = vbTab Then
        ShouldSkipDueToRightCell = True
    Else
        ShouldSkipDueToRightCell = False
    End If
End Function

Private Function ExtractNumericPrefix(ByVal cellValue As String, ByVal allowDots As Boolean, ByRef normalizedValue As String, ByRef remainderValue As String) As Boolean
    Dim workingValue As String
    Dim charValue As String
    Dim index As Long
    
    workingValue = cellValue
    
    Do While Len(workingValue) > 0 And Left$(workingValue, 1) = "'"
        workingValue = Mid$(workingValue, 2)
    Loop
    
    index = 1
    Do While index <= Len(workingValue)
        charValue = Mid$(workingValue, index, 1)
        If charValue = " " Or charValue = ChrW$(&H3000) Or charValue = vbTab Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    workingValue = Mid$(workingValue, index)
    
    normalizedValue = workingValue
    If Len(workingValue) = 0 Then
        remainderValue = ""
        ExtractNumericPrefix = False
        Exit Function
    End If
    
    charValue = Left$(workingValue, 1)
    If charValue < "0" Or charValue > "9" Then
        remainderValue = workingValue
        ExtractNumericPrefix = False
        Exit Function
    End If
    
    index = 1
    Do While index <= Len(workingValue)
        charValue = Mid$(workingValue, index, 1)
        If charValue >= "0" And charValue <= "9" Then
            index = index + 1
        ElseIf allowDots And charValue = "." Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    
    remainderValue = Mid$(workingValue, index)
    ExtractNumericPrefix = (index > 1)
End Function

Private Function HasNumericPrefix(ByVal cellValue As String, ByVal allowDots As Boolean) As Boolean
    Dim normalized As String
    Dim remainder As String
    Dim firstChar As String
    
    If Not ExtractNumericPrefix(cellValue, allowDots, normalized, remainder) Then
        HasNumericPrefix = False
        Exit Function
    End If
    
    If Len(remainder) = 0 Then
        HasNumericPrefix = True
        Exit Function
    End If
    
    firstChar = Left$(remainder, 1)
    If firstChar = " " Or firstChar = ChrW$(&H3000) Or firstChar = vbTab Then
        HasNumericPrefix = True
    Else
        HasNumericPrefix = False
    End If
End Function

Private Function HasCellValue(ByVal targetCell As Range) As Boolean
    Dim cellValue As Variant
    cellValue = targetCell.value
    
    If IsError(cellValue) Then
        HasCellValue = False
        Exit Function
    End If
    
    Dim textValue As String
    textValue = CStr(cellValue)
    textValue = Replace$(textValue, " ", "")
    textValue = Replace$(textValue, ChrW$(&H3000), "")
    
    HasCellValue = Len(textValue) > 0
End Function

Private Sub WriteNormalizedNumber(ByVal targetCell As Range, ByVal newNumber As String, ByVal allowDots As Boolean)
    Dim cellFormula As String
    Dim hasApostrophe As Boolean
    
    On Error Resume Next
    cellFormula = targetCell.Formula
    On Error GoTo 0
    
    hasApostrophe = (Len(cellFormula) > 0 And Left$(cellFormula, 1) = "'")
    If hasApostrophe Then
        cellFormula = Mid$(cellFormula, 2)
    End If
    
    If Len(cellFormula) = 0 Then
        cellFormula = targetCell.Text
    End If
    If Len(cellFormula) = 0 Then
        cellFormula = CStr(targetCell.value)
    End If
    
    Dim newContent As String
    newContent = BuildNumberString(cellFormula, newNumber, allowDots)
    
    targetCell.NumberFormat = "@"
    If hasApostrophe Then
        targetCell.Formula = "'" & newContent
    Else
        targetCell.value = newContent
    End If
End Sub

Private Function BuildNumberString(ByVal originalValue As String, ByVal newNumber As String, ByVal allowDots As Boolean) As String
    Dim workingValue As String
    Dim leadingSpaces As String
    Dim index As Long
    Dim charValue As String
    
    workingValue = originalValue
    
    Do While Len(workingValue) > 0 And Left$(workingValue, 1) = "'"
        workingValue = Mid$(workingValue, 2)
    Loop
    
    index = 1
    Do While index <= Len(workingValue)
        charValue = Mid$(workingValue, index, 1)
        If charValue = " " Or charValue = ChrW$(&H3000) Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    
    leadingSpaces = Left$(workingValue, index - 1)
    workingValue = Mid$(workingValue, index)
    
    index = 1
    Do While index <= Len(workingValue)
        charValue = Mid$(workingValue, index, 1)
        If charValue >= "0" And charValue <= "9" Then
            index = index + 1
        ElseIf allowDots And charValue = "." Then
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    
    Dim suffix As String
    Dim consumedNumber As Boolean
    consumedNumber = (index > 1)
    suffix = Mid$(workingValue, index)
    If Len(suffix) > 0 And Not consumedNumber Then
        charValue = Left$(suffix, 1)
        If charValue <> " " And charValue <> ChrW$(&H3000) Then
            suffix = " " & suffix
        End If
    End If
    
    BuildNumberString = leadingSpaces & newNumber & suffix
End Function




