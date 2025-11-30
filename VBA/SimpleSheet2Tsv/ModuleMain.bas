Attribute VB_Name = "ModuleMain"
Option Explicit

Private Const VER = "1.0.0"


Public Sub RunConversion()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets("main")

    wsMain.Range("A14").ClearContents

    Dim excelPath As String
    Dim outputDir As String
    Dim includeHiddenValue As String

    excelPath = Trim$(wsMain.Range("B7").Value)
    outputDir = Trim$(wsMain.Range("B8").Value)
    includeHiddenValue = UCase$(Trim$(wsMain.Range("B9").Value))

    Dim validationMessage As String
    validationMessage = ValidateInputs(excelPath, outputDir, includeHiddenValue)
    If Len(validationMessage) > 0 Then
        MsgBox validationMessage, vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim includeHidden As Boolean
    includeHidden = (includeHiddenValue = "YES")

    Dim screenUpdating As Boolean
    Dim enableEvents As Boolean
    Dim displayAlerts As Boolean

    screenUpdating = Application.screenUpdating
    enableEvents = Application.enableEvents
    displayAlerts = Application.displayAlerts

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.displayAlerts = False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(fileName:=excelPath, ReadOnly:=True, AddToMru:=False)

    Dim ws As Worksheet
    For Each ws In srcWb.Worksheets
        If includeHidden Or ws.Visible = xlSheetVisible Then
            ExportSheetToTsv ws, outputDir
        End If
    Next ws

    srcWb.Close SaveChanges:=False

Cleanup:
    Application.screenUpdating = screenUpdating
    Application.enableEvents = enableEvents
    Application.displayAlerts = displayAlerts
    Exit Sub

ErrHandler:
    wsMain.Range("A14").Value = "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then
        srcWb.Close SaveChanges:=False
    End If
    GoTo Cleanup
End Sub

Private Function ValidateInputs(ByVal excelPath As String, ByVal outputDir As String, ByVal includeHiddenValue As String) As String
    If Len(excelPath) = 0 Then
        ValidateInputs = "Excelファイルパスが未入力です。"
        Exit Function
    End If

    If Dir(excelPath, vbNormal) = "" Then
        ValidateInputs = "Excelファイルが存在しません。"
        Exit Function
    End If

    Dim lowerPath As String
    lowerPath = LCase$(excelPath)
    If Not (Right$(lowerPath, 5) = ".xlsx" Or Right$(lowerPath, 5) = ".xlsm") Then
        ValidateInputs = "Excelファイルパスは.xlsx または .xlsm を指定してください。"
        Exit Function
    End If

    If Len(outputDir) = 0 Then
        ValidateInputs = "TSV出力フォルダパスが未入力です。"
        Exit Function
    End If

    If Dir(outputDir, vbDirectory) = "" Then
        ValidateInputs = "TSV出力フォルダが存在しません。"
        Exit Function
    End If

    If includeHiddenValue <> "YES" And includeHiddenValue <> "NO" Then
        ValidateInputs = "非表示シートも対象には ""YES"" または ""NO"" を指定してください。"
        Exit Function
    End If
End Function

Private Sub ExportSheetToTsv(ByVal ws As Worksheet, ByVal outputDir As String)
    Dim targetPath As String
    targetPath = BuildOutputPath(outputDir, SanitizeFileName(ws.Name) & ".tsv")

    Dim rng As Range
    Set rng = ws.UsedRange

    Dim data As Variant
    data = rng.Value

    Dim fh As Integer
    fh = FreeFile

    Open targetPath For Output As #fh

    If IsArray(data) Then
        Dim rowCount As Long
        Dim colCount As Long
        rowCount = UBound(data, 1)
        colCount = UBound(data, 2)

        Dim r As Long
        Dim c As Long
        Dim lineParts() As String
        For r = 1 To rowCount
            ReDim lineParts(1 To colCount)
            For c = 1 To colCount
                lineParts(c) = CellToText(data(r, c))
            Next c
            Print #fh, Join(lineParts, vbTab)
        Next r
    Else
        Print #fh, CellToText(data)
    End If

    Close #fh
End Sub

Private Function BuildOutputPath(ByVal folderPath As String, ByVal fileName As String) As String
    If Len(folderPath) = 0 Then
        BuildOutputPath = fileName
    ElseIf Right$(folderPath, 1) = "\" Then
        BuildOutputPath = folderPath & fileName
    Else
        BuildOutputPath = folderPath & "\" & fileName
    End If
End Function

Private Function SanitizeFileName(ByVal sheetName As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim i As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        sheetName = Replace(sheetName, invalidChars(i), "_")
    Next i

    If Len(sheetName) = 0 Then
        sheetName = "sheet"
    End If

    SanitizeFileName = sheetName
End Function

Private Function CellToText(ByVal cellValue As Variant) As String
    Dim result As String
    On Error Resume Next
    result = CStr(cellValue)
    If Err.Number <> 0 Then
        result = "#ERROR"
    End If
    On Error GoTo 0
    CellToText = result
End Function

