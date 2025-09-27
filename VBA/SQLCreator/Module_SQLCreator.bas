Option Explicit

Private Const CATEGORY_NUMERIC As String = "数値"
Private Const CATEGORY_STRING As String = "文字列"
Private Const CATEGORY_DATE As String = "日付"
Private Const CATEGORY_TIME As String = "時刻"
Private Const CATEGORY_TIMESTAMP As String = "日時"

Private Type ColumnDefinition
    Name As String
    DataType As String
    category As String
    Precision As Long
    HasPrecision As Boolean
    Scale As Long
    HasScale As Boolean
    IsPrimaryKey As Boolean
    IsNotNull As Boolean
End Type

Public Sub ボタン1_Click()
    ProcessInstructionRows
End Sub

Private Sub ProcessInstructionRows()
    On Error GoTo FatalError
    Dim mainWs As Worksheet
    Dim typeWs As Worksheet
    Dim typeCatalog As Object
    Dim currentRow As Long
    Dim targetMark As String

    Set mainWs = ThisWorkbook.Worksheets("main")
    Set typeWs = ThisWorkbook.Worksheets("type")
    Set typeCatalog = BuildTypeCatalog(typeWs)

    If typeCatalog Is Nothing Or typeCatalog.Count = 0 Then
        MsgBox "typeシートの型定義が取得できません。", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    currentRow = 11
    Do While LenB(Trim$(CStr(mainWs.Cells(currentRow, 2).Value))) > 0
        targetMark = Trim$(CStr(mainWs.Cells(currentRow, 3).Value))
        mainWs.Cells(currentRow, 15).Value = vbNullString
        If targetMark = "○" Then
            ProcessSingleInstruction mainWs, typeCatalog, currentRow
        End If
        currentRow = currentRow + 1
    Loop

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

FatalError:
    MsgBox "SQL作成中に致命的なエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub ProcessSingleInstruction(ByVal mainWs As Worksheet, ByVal typeCatalog As Object, ByVal rowIndex As Long)
    Dim errors As Collection
    Dim filePath As String
    Dim sheetName As String
    Dim tableName As String
    Dim colNameAddr As String
    Dim typeAddr As String
    Dim precisionAddr As String
    Dim scaleAddr As String
    Dim pkAddr As String
    Dim notNullAddr As String
    Dim dbms As String
    Dim definitionWb As Workbook
    Dim definitionWs As Worksheet
    Dim columns As Collection
    Dim sqlText As String
    Dim outputPath As String

    Set errors = New Collection

    filePath = Trim$(CStr(mainWs.Cells(rowIndex, 4).Value))
    sheetName = Trim$(CStr(mainWs.Cells(rowIndex, 5).Value))
    tableName = Trim$(CStr(mainWs.Cells(rowIndex, 6).Value))
    colNameAddr = Trim$(CStr(mainWs.Cells(rowIndex, 7).Value))
    typeAddr = Trim$(CStr(mainWs.Cells(rowIndex, 8).Value))
    precisionAddr = Trim$(CStr(mainWs.Cells(rowIndex, 9).Value))
    scaleAddr = Trim$(CStr(mainWs.Cells(rowIndex, 10).Value))
    pkAddr = Trim$(CStr(mainWs.Cells(rowIndex, 11).Value))
    notNullAddr = Trim$(CStr(mainWs.Cells(rowIndex, 12).Value))
    dbms = Trim$(CStr(mainWs.Cells(rowIndex, 13).Value))

    ValidateRequiredValue filePath, "ファイルパス", errors
    ValidateRequiredValue sheetName, "シート名", errors
    ValidateRequiredValue tableName, "テーブル名", errors
    ValidateRequiredValue colNameAddr, "カラム名開始セル", errors
    ValidateRequiredValue typeAddr, "型カラム開始セル", errors
    ValidateRequiredValue precisionAddr, "整数桁カラム開始セル", errors
    ValidateRequiredValue scaleAddr, "小数桁カラム開始セル", errors
    ValidateRequiredValue pkAddr, "PKカラム開始セル", errors
    ValidateRequiredValue notNullAddr, "NotNullカラム開始セル", errors
    ValidateDbms dbms, errors

    If errors.Count > 0 Then
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If Dir$(filePath, vbNormal) = vbNullString Then
        AddError errors, "ファイルパスが不正です: " & filePath
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    On Error Resume Next
    Set definitionWb = Application.Workbooks.Open(fileName:=filePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        AddError errors, "Excelファイルを開けません: " & filePath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    Set definitionWs = definitionWb.Worksheets(sheetName)
    On Error GoTo 0
    If definitionWs Is Nothing Then
        AddError errors, "指定シートが見つかりません: " & sheetName
        SafeCloseWorkbook definitionWb
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    Set columns = ReadColumnDefinitions(definitionWs, colNameAddr, typeAddr, precisionAddr, scaleAddr, pkAddr, notNullAddr, typeCatalog, errors)

    SafeCloseWorkbook definitionWb

    If errors.Count > 0 Then
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "有効なカラム定義が1件もありません。"
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    sqlText = GenerateSqlText(tableName, columns, dbms, errors)
    If errors.Count > 0 Then
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    outputPath = WriteSqlFile(filePath, tableName, sqlText, errors)
    If errors.Count > 0 Then
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    mainWs.Cells(rowIndex, 15).Value = "SQL出力: " & outputPath
End Sub

Private Function ReadColumnDefinitions(ByVal targetWs As Worksheet, ByVal nameAddr As String, _
                                       ByVal typeAddr As String, ByVal precisionAddr As String, _
                                       ByVal scaleAddr As String, ByVal pkAddr As String, _
                                       ByVal notNullAddr As String, ByVal typeCatalog As Object, _
                                       ByVal errors As Collection) As Collection

    Dim nameCell As Range
    Dim typeCell As Range
    Dim precisionCell As Range
    Dim scaleCell As Range
    Dim pkCell As Range
    Dim notNullCell As Range

    Set nameCell = ResolveSingleCell(targetWs, nameAddr, "カラム名開始セル", errors)
    Set typeCell = ResolveSingleCell(targetWs, typeAddr, "型カラム開始セル", errors)
    Set precisionCell = ResolveSingleCell(targetWs, precisionAddr, "整数桁カラム開始セル", errors)
    Set scaleCell = ResolveSingleCell(targetWs, scaleAddr, "小数桁カラム開始セル", errors)
    Set pkCell = ResolveSingleCell(targetWs, pkAddr, "PKカラム開始セル", errors)
    Set notNullCell = ResolveSingleCell(targetWs, notNullAddr, "NotNullカラム開始セル", errors)

    If errors.Count > 0 Then
        Exit Function
    End If

    Dim columns As New Collection
    Dim currentRow As Long
    currentRow = nameCell.Row

    Do While currentRow <= targetWs.Rows.Count
        Dim columnName As String
        columnName = Trim$(CStr(targetWs.Cells(currentRow, nameCell.Column).Value))
        If LenB(columnName) = 0 Then Exit Do

        Dim columnType As String
        columnType = Trim$(CStr(targetWs.Cells(currentRow, typeCell.Column).Value))
        If LenB(columnType) = 0 Then
            AddError errors, "型が未入力です: " & targetWs.Cells(currentRow, typeCell.Column).Address(False, False)
            Exit Do
        End If

        Dim typeKey As String
        typeKey = UCase$(columnType)
        If Not typeCatalog.Exists(typeKey) Then
            AddError errors, "型がtypeシートに存在しません: " & columnType & " (" & targetWs.Cells(currentRow, typeCell.Column).Address(False, False) & ")"
            Exit Do
        End If

        Dim category As String
        category = CStr(typeCatalog(typeKey))

        Dim definition As ColumnDefinition
        definition.Name = columnName
        definition.DataType = columnType
        definition.category = category

        Dim precisionText As String
        precisionText = Trim$(CStr(targetWs.Cells(currentRow, precisionCell.Column).Value))
        Dim scaleText As String
        scaleText = Trim$(CStr(targetWs.Cells(currentRow, scaleCell.Column).Value))

        If category = CATEGORY_NUMERIC Or category = CATEGORY_STRING Then
            If Not TryParsePositiveInteger(precisionText, definition.Precision) Then
                AddError errors, "整数桁が不正です: " & targetWs.Cells(currentRow, precisionCell.Column).Address(False, False)
            Else
                definition.HasPrecision = True
            End If
        End If

        If category = CATEGORY_NUMERIC Then
            If Not TryParseNonNegativeInteger(scaleText, definition.Scale) Then
                AddError errors, "小数桁が不正です: " & targetWs.Cells(currentRow, scaleCell.Column).Address(False, False)
            Else
                definition.HasScale = True
            End If
        End If

        If errors.Count > 0 Then
            Exit Do
        End If

        definition.IsPrimaryKey = (LenB(Trim$(CStr(targetWs.Cells(currentRow, pkCell.Column).Value))) > 0)
        definition.IsNotNull = (LenB(Trim$(CStr(targetWs.Cells(currentRow, notNullCell.Column).Value))) > 0) Or definition.IsPrimaryKey

        columns.Add definition
        currentRow = currentRow + 1
    Loop

    If errors.Count = 0 Then
        Set ReadColumnDefinitions = columns
    End If
End Function

Private Function GenerateSqlText(ByVal tableName As String, ByVal columns As Collection, ByVal dbms As String, ByVal errors As Collection) As String
    Dim schemaName As String
    Dim pureTableName As String
    Dim qualifiedName As String

    schemaName = ExtractSchemaName(tableName)
    pureTableName = ExtractTableName(tableName)

    Select Case dbms
        Case "SQLServer"
            If LenB(schemaName) = 0 Then schemaName = "dbo"
        Case "SQLite"
            schemaName = vbNullString
    End Select

    qualifiedName = BuildQualifiedName(schemaName, pureTableName, dbms)

    Dim lines() As String
    ReDim lines(0 To columns.Count)
    Dim idx As Long
    Dim pkColumns As New Collection

    For idx = 1 To columns.Count
        Dim def As ColumnDefinition
        def = columns(idx)
        Dim lineText As String
        lineText = "    " & QuoteIdentifier(def.Name, dbms) & " " & FormatDataType(def, dbms, errors)
        If errors.Count > 0 Then Exit Function
        If def.IsNotNull Then
            lineText = lineText & " NOT NULL"
        End If
        lines(idx - 1) = lineText
        If def.IsPrimaryKey Then
            pkColumns.Add QuoteIdentifier(def.Name, dbms)
        End If
    Next idx

    Dim constraintLine As String
    If pkColumns.Count > 0 Then
        constraintLine = "    CONSTRAINT " & QuoteIdentifier("PK_" & SanitizeForConstraint(pureTableName), dbms) & _
                         " PRIMARY KEY (" & JoinCollection(pkColumns, ", ") & ")"
        lines(columns.Count) = constraintLine
    Else
        ReDim Preserve lines(0 To columns.Count - 1)
    End If

    Dim dropSql As String
    dropSql = BuildDropStatement(qualifiedName, dbms)

    Dim createSql As String
    createSql = "CREATE TABLE " & qualifiedName & vbCrLf & "(" & vbCrLf & Join(lines, "," & vbCrLf) & vbCrLf & ");"

    GenerateSqlText = dropSql & vbCrLf & vbCrLf & createSql
End Function

Private Function FormatDataType(ByRef definition As ColumnDefinition, ByVal dbms As String, ByVal errors As Collection) As String
    Dim typeText As String
    typeText = definition.DataType

    Select Case definition.category
        Case CATEGORY_NUMERIC
            If Not definition.HasPrecision Or Not definition.HasScale Then
                AddError errors, "数値型の桁情報が不足しています: " & definition.Name
                Exit Function
            End If
            FormatDataType = typeText & "(" & CStr(definition.Precision) & "," & CStr(definition.Scale) & ")"
        Case CATEGORY_STRING
            If Not definition.HasPrecision Then
                AddError errors, "文字列型の桁情報が不足しています: " & definition.Name
                Exit Function
            End If
            FormatDataType = typeText & "(" & CStr(definition.Precision) & ")"
        Case Else
            FormatDataType = typeText
    End Select
End Function

Private Function BuildDropStatement(ByVal qualifiedName As String, ByVal dbms As String) As String
    Select Case dbms
        Case "SQLServer"
            BuildDropStatement = "DROP TABLE IF EXISTS " & qualifiedName & ";"
        Case "SQLite"
            BuildDropStatement = "DROP TABLE IF EXISTS " & qualifiedName & ";"
        Case Else
            BuildDropStatement = "DROP TABLE " & qualifiedName & ";"
    End Select
End Function

Private Function WriteSqlFile(ByVal sourceFilePath As String, ByVal tableName As String, _
                              ByVal sqlText As String, ByVal errors As Collection) As String
    Dim fso As Object
    Dim parentFolder As String
    Dim fileName As String
    Dim outputPath As String
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(sourceFilePath)
    If LenB(parentFolder) = 0 Then
        AddError errors, "テーブル定義ファイルの親フォルダを取得できません: " & sourceFilePath
        Exit Function
    End If

    fileName = Format$(Now, "yyyymmdd_hhnnss") & "_" & SanitizeForFile(tableName:=tableName) & ".sql"
    outputPath = fso.BuildPath(parentFolder, fileName)

    On Error Resume Next
    Set ts = fso.CreateTextFile(outputPath, True, False)
    If Err.Number <> 0 Then
        AddError errors, "SQLファイルを作成できません: " & outputPath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ts.Write sqlText
    ts.Close

    WriteSqlFile = outputPath
End Function

Private Function BuildTypeCatalog(ByVal typeWs As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim currentRow As Long
    currentRow = 3

    Do While LenB(Trim$(CStr(typeWs.Cells(currentRow, 2).Value))) > 0
        Dim category As String
        category = Trim$(CStr(typeWs.Cells(currentRow, 2).Value))
        Dim currentColumn As Long
        currentColumn = 3
        Do While LenB(Trim$(CStr(typeWs.Cells(currentRow, currentColumn).Value))) > 0
            Dim typeName As String
            typeName = UCase$(Trim$(CStr(typeWs.Cells(currentRow, currentColumn).Value)))
            If LenB(typeName) > 0 Then
                If Not dict.Exists(typeName) Then
                    dict.Add typeName, category
                End If
            End If
            currentColumn = currentColumn + 1
        Loop
        currentRow = currentRow + 1
    Loop

    Set BuildTypeCatalog = dict
End Function

Private Function ResolveSingleCell(ByVal targetWs As Worksheet, ByVal addressText As String, _
                                   ByVal label As String, ByVal errors As Collection) As Range
    On Error Resume Next
    Set ResolveSingleCell = targetWs.Range(addressText)
    If Err.Number <> 0 Then
        AddError errors, label & " のセル参照が不正です: " & addressText
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub ValidateRequiredValue(ByVal valueText As String, ByVal label As String, ByVal errors As Collection)
    If LenB(valueText) = 0 Then
        AddError errors, label & " が未入力です。"
    End If
End Sub

Private Sub ValidateDbms(ByVal dbms As String, ByVal errors As Collection)
    Select Case UCase$(dbms)
        Case "ORACLE", "SQLSERVER", "SQLITE"
            ' OK
        Case Else
            AddError errors, "DBMSが不正です: " & dbms
    End Select
End Sub

Private Function ExtractSchemaName(ByVal tableName As String) As String
    Dim parts() As String
    parts = Split(tableName, ".")
    If UBound(parts) >= 1 Then
        ExtractSchemaName = Trim$(parts(0))
    Else
        ExtractSchemaName = vbNullString
    End If
End Function

Private Function ExtractTableName(ByVal tableName As String) As String
    Dim parts() As String
    parts = Split(tableName, ".")
    ExtractTableName = Trim$(parts(UBound(parts)))
End Function

Private Function BuildQualifiedName(ByVal schemaName As String, ByVal tableName As String, ByVal dbms As String) As String
    If LenB(schemaName) > 0 Then
        BuildQualifiedName = QuoteIdentifier(schemaName, dbms) & "." & QuoteIdentifier(tableName, dbms)
    Else
        BuildQualifiedName = QuoteIdentifier(tableName, dbms)
    End If
End Function

Private Function QuoteIdentifier(ByVal identifier As String, ByVal dbms As String) As String
    Dim cleaned As String
    cleaned = Trim$(identifier)
    Select Case dbms
        Case "SQLServer"
            cleaned = Replace(cleaned, "[", vbNullString)
            cleaned = Replace(cleaned, "]", vbNullString)
            QuoteIdentifier = "[" & cleaned & "]"
        Case Else
            cleaned = Replace(cleaned, """", vbNullString)
            QuoteIdentifier = """" & cleaned & """"
    End Select
End Function


Private Function SanitizeForConstraint(ByVal tableName As String) As String
    Dim cleaned As String
    cleaned = Trim$(tableName)
    cleaned = Replace(cleaned, "[", vbNullString)
    cleaned = Replace(cleaned, "]", vbNullString)
    cleaned = Replace(cleaned, """", vbNullString)
    cleaned = Replace(cleaned, ".", "_")
    cleaned = Replace(cleaned, " ", "_")
    cleaned = Replace(cleaned, "-", "_")
    SanitizeForConstraint = cleaned
End Function



Private Function SanitizeForFile(ByVal tableName As String) As String
    Dim cleaned As String
    cleaned = Trim$(tableName)
    cleaned = Replace(cleaned, "[", vbNullString)
    cleaned = Replace(cleaned, "]", vbNullString)
    cleaned = Replace(cleaned, """", vbNullString)
    cleaned = Replace(cleaned, ":", "_")
    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "\", "_")
    cleaned = Replace(cleaned, "*", "_")
    cleaned = Replace(cleaned, "?", "_")
    cleaned = Replace(cleaned, "<", "_")
    cleaned = Replace(cleaned, ">", "_")
    cleaned = Replace(cleaned, "|", "_")
    cleaned = Replace(cleaned, " ", "_")
    cleaned = Replace(cleaned, ".", "_")
    SanitizeForFile = cleaned
End Function


Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    If items Is Nothing Then Exit Function
    If items.Count = 0 Then Exit Function
    Dim parts() As String
    Dim i As Long
    ReDim parts(1 To items.Count)
    For i = 1 To items.Count
        parts(i) = CStr(items(i))
    Next i
    JoinCollection = Join(parts, delimiter)
End Function

Private Sub AddError(ByVal errors As Collection, ByVal message As String)
    errors.Add message
End Sub

Private Sub WriteErrors(ByVal mainWs As Worksheet, ByVal rowIndex As Long, ByVal errors As Collection)
    mainWs.Cells(rowIndex, 15).Value = JoinCollection(errors, vbLf)
End Sub

Private Function TryParsePositiveInteger(ByVal textValue As String, ByRef resultValue As Long) As Boolean
    If LenB(textValue) = 0 Then Exit Function
    If Not IsNumeric(textValue) Then Exit Function
    Dim temp As Double
    temp = CDbl(textValue)
    If temp <= 0# Then Exit Function
    If temp <> Fix(temp) Then Exit Function
    resultValue = CLng(temp)
    TryParsePositiveInteger = True
End Function

Private Function TryParseNonNegativeInteger(ByVal textValue As String, ByRef resultValue As Long) As Boolean
    If LenB(textValue) = 0 Then Exit Function
    If Not IsNumeric(textValue) Then Exit Function
    Dim temp As Double
    temp = CDbl(textValue)
    If temp < 0# Then Exit Function
    If temp <> Fix(temp) Then Exit Function
    resultValue = CLng(temp)
    TryParseNonNegativeInteger = True
End Function

Private Sub SafeCloseWorkbook(ByVal targetWb As Workbook)
    On Error Resume Next
    If Not targetWb Is Nothing Then
        targetWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub


