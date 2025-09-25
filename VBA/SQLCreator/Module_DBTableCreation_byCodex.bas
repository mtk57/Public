Option Explicit

Public Sub DBテーブル作成(ByVal tsvFilePath As String)
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim inputFile As Object
    Dim lines As Collection
    Dim lineText As String
    Dim tableNameRaw As String
    Dim schemaName As String
    Dim tableName As String
    Dim qualifiedTableName As String
    Dim outputSql As String
    Dim columnLines As Collection
    Dim pkColumns As Collection
    Dim lineIndex As Long
    Dim headerValues As Variant
    Dim columnValues As Variant
    Dim columnLine As String
    Dim columnName As String
    Dim columnType As String
    Dim columnLength As String
    Dim pkFlag As String
    Dim parentFolder As String
    Dim outputPath As String
    Dim outputFile As Object

    If LenB(tsvFilePath) = 0 Then
        Err.Raise vbObjectError + 1000, "DBテーブル作成", "TSVファイルパスが指定されていません。"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(tsvFilePath) Then
        Err.Raise vbObjectError + 1001, "DBテーブル作成", "指定したTSVファイルが見つかりません: " & tsvFilePath
    End If

    Set lines = New Collection
    Set inputFile = fso.OpenTextFile(tsvFilePath, 1, False, -2)
    Do While Not inputFile.AtEndOfStream
        lineText = inputFile.ReadLine
        lines.Add lineText
    Loop
    inputFile.Close
    Set inputFile = Nothing

    If lines.Count < 3 Then
        Err.Raise vbObjectError + 1002, "DBテーブル作成", "TSVファイルの行数が不足しています。(最小3行必要)"
    End If

    tableNameRaw = Trim$(CStr(lines.Item(1)))
    If LenB(tableNameRaw) = 0 Then
        Err.Raise vbObjectError + 1003, "DBテーブル作成", "テーブル名が空です。"
    End If

    schemaName = ExtractSchemaName(tableNameRaw)
    tableName = ExtractTableName(tableNameRaw)
    qualifiedTableName = WrapWithBrackets(schemaName) & "." & WrapWithBrackets(tableName)

    headerValues = ParseTsvLine(CStr(lines.Item(2)), 4)
    ValidateHeader headerValues

    Set columnLines = New Collection
    Set pkColumns = New Collection

    For lineIndex = 3 To lines.Count
        lineText = Trim$(CStr(lines.Item(lineIndex)))
        If LenB(lineText) = 0 Then
            ' Skip empty lines
            GoTo ContinueLoop
        End If

        columnValues = ParseTsvLine(lineText, 4)
        columnName = Trim$(CStr(columnValues(0)))
        columnType = Trim$(CStr(columnValues(1)))
        columnLength = Trim$(CStr(columnValues(2)))
        pkFlag = UCase$(Trim$(CStr(columnValues(3))))

        ValidateColumnDefinition lineIndex, columnName, columnType

        columnLine = "    " & WrapWithBrackets(columnName) & " " & columnType
        If ShouldAppendLength(columnType, columnLength) Then
            columnLine = columnLine & FormatLengthSpecifier(columnLength)
        End If
        columnLines.Add columnLine

        If pkFlag = "Y" Then
            pkColumns.Add WrapWithBrackets(columnName)
        End If
ContinueLoop:
    Next lineIndex

    If columnLines.Count = 0 Then
        Err.Raise vbObjectError + 1004, "DBテーブル作成", "有効なカラム定義が1件もありません。"
    End If

    outputSql = BuildCreateTableSql(qualifiedTableName, columnLines, pkColumns, tableName)

    parentFolder = fso.GetParentFolderName(tsvFilePath)
    If LenB(parentFolder) = 0 Then
        Err.Raise vbObjectError + 1005, "DBテーブル作成", "TSVファイルのフォルダパスを取得できません。"
    End If

    outputPath = fso.BuildPath(parentFolder, BuildTimestamp() & ".sql")
    Set outputFile = fso.CreateTextFile(outputPath, True, False)
    outputFile.Write outputSql
    outputFile.Close
    Set outputFile = Nothing

    Debug.Print "DBテーブル作成: SQLファイルを出力しました => " & outputPath
    Exit Sub

ErrorHandler:
    Dim errMsg As String
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not inputFile Is Nothing Then inputFile.Close
    If Not outputFile Is Nothing Then outputFile.Close
    On Error GoTo 0
    errMsg = "DBテーブル作成でエラーが発生しました。" & vbCrLf & _
             "TSVファイル: " & tsvFilePath & vbCrLf & _
             "エラー番号: " & errNumber & vbCrLf & _
             "詳細: " & errDescription
    Debug.Print errMsg
    MsgBox errMsg, vbCritical, "DBテーブル作成"
End Sub

Private Function ParseTsvLine(ByVal lineText As String, ByVal expectedColumns As Long) As Variant
    Dim rawValues() As String
    Dim normalizedValues() As String
    Dim i As Long

    rawValues = Split(lineText, vbTab)
    ReDim normalizedValues(0 To expectedColumns - 1)

    For i = 0 To expectedColumns - 1
        If i <= UBound(rawValues) Then
            normalizedValues(i) = rawValues(i)
        Else
            normalizedValues(i) = vbNullString
        End If
    Next i

    ParseTsvLine = normalizedValues
End Function

Private Sub ValidateHeader(ByVal headerValues As Variant)
    Dim expectedHeaders As Variant
    expectedHeaders = Array("カラム名", "カラム型", "カラム桁数", "PK")

    Dim i As Long
    For i = LBound(expectedHeaders) To UBound(expectedHeaders)
        If StrComp(Trim$(headerValues(i)), expectedHeaders(i), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1100, "DBテーブル作成", _
                "TSVヘッダーの" & (i + 1) & "列目が想定と異なります。期待値: " & expectedHeaders(i)
        End If
    Next i
End Sub

Private Sub ValidateColumnDefinition(ByVal lineIndex As Long, ByVal columnName As String, ByVal columnType As String)
    If LenB(columnName) = 0 Then
        Err.Raise vbObjectError + 1200, "DBテーブル作成", CStr(lineIndex) & "行目: カラム名が空です。"
    End If
    If LenB(columnType) = 0 Then
        Err.Raise vbObjectError + 1201, "DBテーブル作成", CStr(lineIndex) & "行目: カラム型が空です。"
    End If
End Sub

Private Function ShouldAppendLength(ByVal columnType As String, ByVal columnLength As String) As Boolean
    If LenB(columnLength) = 0 Then
        ShouldAppendLength = False
        Exit Function
    End If

    If columnLength = "-" Then
        ShouldAppendLength = False
        Exit Function
    End If

    ShouldAppendLength = True
End Function

Private Function FormatLengthSpecifier(ByVal columnLength As String) As String
    FormatLengthSpecifier = "(" & columnLength & ")"
End Function

Private Function BuildCreateTableSql(ByVal qualifiedTableName As String, _
                                     ByVal columnLines As Collection, _
                                     ByVal pkColumns As Collection, _
                                     ByVal rawTableName As String) As String
    Dim sqlText As String
    Dim i As Long
    sqlText = "CREATE TABLE " & qualifiedTableName & vbCrLf & "(" & vbCrLf

    For i = 1 To columnLines.Count
        sqlText = sqlText & columnLines.Item(i)
        If i < columnLines.Count Or pkColumns.Count > 0 Then
            sqlText = sqlText & ","
        End If
        sqlText = sqlText & vbCrLf
    Next i

    If pkColumns.Count > 0 Then
        sqlText = sqlText & "    CONSTRAINT " & WrapWithBrackets("PK_" & SanitizeForConstraint(rawTableName)) & _
                  " PRIMARY KEY (" & JoinCollection(pkColumns, ", ") & ")" & vbCrLf
    End If

    sqlText = sqlText & ");" & vbCrLf
    BuildCreateTableSql = sqlText
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim temp() As String
    Dim i As Long

    ReDim temp(0 To items.Count - 1)
    For i = 1 To items.Count
        temp(i - 1) = items.Item(i)
    Next i

    JoinCollection = Join(temp, delimiter)
End Function

Private Function ExtractSchemaName(ByVal rawTableName As String) As String
    Dim dotPosition As Long
    Dim schemaName As String

    dotPosition = InStr(rawTableName, ".")
    If dotPosition > 0 Then
        schemaName = Trim$(Left$(rawTableName, dotPosition - 1))
    Else
        schemaName = "dbo"
    End If

    If LenB(schemaName) = 0 Then
        schemaName = "dbo"
    End If

    ExtractSchemaName = RemoveBrackets(schemaName)
End Function

Private Function ExtractTableName(ByVal rawTableName As String) As String
    Dim dotPosition As Long
    Dim tableName As String

    dotPosition = InStr(rawTableName, ".")
    If dotPosition > 0 Then
        tableName = Trim$(Mid$(rawTableName, dotPosition + 1))
    Else
        tableName = Trim$(rawTableName)
    End If

    If LenB(tableName) = 0 Then
        Err.Raise vbObjectError + 1300, "DBテーブル作成", "テーブル名を解析できません。"
    End If

    ExtractTableName = RemoveBrackets(tableName)
End Function

Private Function WrapWithBrackets(ByVal identifier As String) As String
    WrapWithBrackets = "[" & RemoveBrackets(identifier) & "]"
End Function

Private Function RemoveBrackets(ByVal identifier As String) As String
    Dim trimmedValue As String
    trimmedValue = Trim$(identifier)
    trimmedValue = Replace(trimmedValue, "[", vbNullString)
    trimmedValue = Replace(trimmedValue, "]", vbNullString)
    RemoveBrackets = trimmedValue
End Function

Private Function SanitizeForConstraint(ByVal tableName As String) As String
    Dim sanitized As String
    sanitized = RemoveBrackets(tableName)
    sanitized = Replace(sanitized, " ", "_")
    sanitized = Replace(sanitized, ".", "_")
    SanitizeForConstraint = sanitized
End Function

Private Function BuildTimestamp() As String
    Dim nowValue As Date
    nowValue = Now
    BuildTimestamp = Format$(nowValue, "yyyymmdd") & "_" & _
                     Right$("0" & CStr(Hour(nowValue)), 2) & _
                     Right$("0" & CStr(Minute(nowValue)), 2) & _
                     Right$("0" & CStr(Second(nowValue)), 2)
End Function
