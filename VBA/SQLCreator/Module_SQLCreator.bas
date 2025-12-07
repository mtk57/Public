Attribute VB_Name = "Module_SQLCreator"
Option Explicit

Private Const VER As String = "2.5.0"

Private Const LOG_ENABLED As Boolean = False
Private Const LOG_FILE_NAME As String = "SQLCreator_debug.log"

Private Const UNSUPPORTED_CHAR_REPLACEMENT As String = "_"
Private Const ENV_DEPENDENT_CHAR_CODES As String = "9312,9313,9314,9315,9316,9317,9318,9319,9320,9321,9322,9323,9324,9325,9326,9327,9328,9329,9330,9331,12964,12965,12966,12967,12968,12969,12970,12971,12972,12973,12974,12975,12976,12953,12849,12850,12857,13182,13181,13180,13179,13129,13130,13076,13080,13090,13091,13094,13095,13110,13115,13116,13127,13128,13143,13133"
Private Const LCID_JAPANESE As Long = &H411

Private logFilePath As String
Private logInitialized As Boolean
Private hasUnsupportedCharReplacement As Boolean

Private Const CATEGORY_NUMERIC As String = "数値"
Private Const CATEGORY_STRING As String = "文字列"
Private Const CATEGORY_DATE As String = "日付"
Private Const CATEGORY_TIME As String = "時刻"
Private Const CATEGORY_TIMESTAMP As String = "日時"

Private Const MAIN_COL_NO As Long = 2
Private Const MAIN_COL_DEF_TARGET As Long = 3
Private Const MAIN_COL_DEF_FILE As Long = 4
Private Const MAIN_COL_DEF_SHEET As Long = 5
Private Const MAIN_COL_DEF_TABLE As Long = 6
Private Const MAIN_COL_COLNAME_ADDR As Long = 7
Private Const MAIN_COL_TYPE_ADDR As Long = 8
Private Const MAIN_COL_PRECISION_ADDR As Long = 9
Private Const MAIN_COL_SCALE_ADDR As Long = 10
Private Const MAIN_COL_PK_ADDR As Long = 11
Private Const MAIN_COL_NOTNULL_ADDR As Long = 12
Private Const MAIN_COL_DBMS As Long = 13
Private Const MAIN_COL_DATA_TARGET As Long = 14
Private Const MAIN_COL_DATA_FILE As Long = 15
Private Const MAIN_COL_DATA_SHEET As Long = 16
Private Const MAIN_COL_DATA_TABLE As Long = 17
Private Const MAIN_COL_DATA_START As Long = 18
Private Const MAIN_COL_DTO_LANG As Long = 19
Private Const MAIN_COL_DTO_CLASS As Long = 20
Private Const MAIN_COL_MESSAGE As Long = 25

Private Const ORACLE_TIME_BASE_DATE As String = "1970-01-01"

Private Sub InitializeLogger()
    If Not LOG_ENABLED Then Exit Sub
    Dim baseFolder As String
    baseFolder = ThisWorkbook.Path
    If LenB(baseFolder) = 0 Then
        On Error Resume Next
        baseFolder = Environ$("TEMP")
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo 0
    End If
    If LenB(baseFolder) = 0 Then Exit Sub

    logFilePath = baseFolder & Application.PathSeparator & LOG_FILE_NAME

    On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(logFilePath, True, False)
    If Err.Number = 0 Then
        ts.WriteLine "==== Log Start (" & Format$(Now, "yyyy-mm-dd HH:nn:ss") & ") ===="
        ts.Close
        logInitialized = True
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureLoggerInitialized()
    If Not LOG_ENABLED Then Exit Sub
    If logInitialized Then Exit Sub
    InitializeLogger
End Sub

Private Sub AppendLogLine(ByVal level As String, ByVal message As String)
    If Not LOG_ENABLED Then Exit Sub
    EnsureLoggerInitialized
    If Not logInitialized Or LenB(logFilePath) = 0 Then Exit Sub

    Dim timestamp As String
    timestamp = Format$(Now, "yyyy-mm-dd HH:nn:ss")

    On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logFilePath, 8, True)
    If Err.Number = 0 Then
        Dim lineText As String
        lineText = "[" & timestamp & "] [" & level & "] " & message
        ts.WriteLine lineText
        ts.Close
        Debug.Print lineText
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub LogInfo(ByVal message As String)
    AppendLogLine "INFO", message
End Sub

Private Sub LogDebug(ByVal message As String)
    AppendLogLine "DEBUG", message
End Sub

Private Sub LogErrorMessage(ByVal message As String)
    AppendLogLine "ERROR", message
End Sub

Private Function DescribeVariant(ByVal value As Variant) As String
    On Error GoTo VariantError
    If IsObject(value) Then
        DescribeVariant = "<Object:" & TypeName(value) & ">"
    ElseIf IsArray(value) Then
        DescribeVariant = "<Array>"
    ElseIf IsError(value) Then
        DescribeVariant = "<Error>"
    ElseIf IsNull(value) Then
        DescribeVariant = "<Null>"
    ElseIf IsEmpty(value) Then
        DescribeVariant = "<Empty>"
    Else
        DescribeVariant = CStr(value) & " (Type=" & TypeName(value) & ")"
    End If
    Exit Function
VariantError:
    DescribeVariant = "<Unprintable>"
End Function

Private Function DescribeStringArray(ByRef values() As String) As String
    On Error GoTo DescribeError
    Dim result As String
    Dim idx As Long
    For idx = LBound(values) To UBound(values)
        If LenB(result) > 0 Then result = result & ", "
        result = result & values(idx)
    Next idx
    DescribeStringArray = result
    Exit Function
DescribeError:
    DescribeStringArray = "<Unavailable>"
End Function

Private Function DescribeVariantArray(ByRef values() As Variant) As String
    On Error GoTo DescribeError
    Dim result As String
    Dim idx As Long
    For idx = LBound(values) To UBound(values)
        If LenB(result) > 0 Then result = result & ", "
        result = result & DescribeVariant(values(idx))
    Next idx
    DescribeVariantArray = result
    Exit Function
DescribeError:
    DescribeVariantArray = "<Unavailable>"
End Function

Private Function ReplaceUnsupportedCharacters(ByVal textValue As String) As String
    Dim idx As Long
    Dim ch As String
    Dim builder As String
    Dim replaced As Boolean

    If LenB(textValue) = 0 Then
        ReplaceUnsupportedCharacters = textValue
        Exit Function
    End If

    For idx = 1 To Len(textValue)
        ch = Mid$(textValue, idx, 1)
        If IsEnvironmentDependentCharacter(ch) Then
            builder = builder & UNSUPPORTED_CHAR_REPLACEMENT
            replaced = True
        Else
            builder = builder & ch
        End If
    Next idx

    If replaced Then
        hasUnsupportedCharReplacement = True
        If LOG_ENABLED Then
            LogDebug "ReplaceUnsupportedCharacters: at least one character was replaced."
        End If
    End If

    ReplaceUnsupportedCharacters = builder
End Function

Private Function IsEnvironmentDependentCharacter(ByVal ch As String) As Boolean
    Static envChars As Object
    If LenB(ch) = 0 Then Exit Function

    If envChars Is Nothing Then
        Set envChars = CreateObject("Scripting.Dictionary")
        Dim codes As Variant
        codes = Split(ENV_DEPENDENT_CHAR_CODES, ",")
        Dim codeText As Variant
        For Each codeText In codes
            codeText = Trim(CStr(codeText))
            If LenB(codeText) > 0 Then
                envChars(ChrW$(CLng(codeText))) = True
            End If
        Next codeText
    End If

    If envChars.Exists(ch) Then
        IsEnvironmentDependentCharacter = True
    Else
        IsEnvironmentDependentCharacter = Not CanRepresentInJapaneseCodePage(ch)
    End If
End Function

Private Function CanRepresentInJapaneseCodePage(ByVal ch As String) As Boolean
    On Error GoTo Failure
    Dim bytes() As Byte
    Dim roundtrip As String

    bytes = StrConv(ch, vbFromUnicode, LCID_JAPANESE)
    roundtrip = StrConv(bytes, vbUnicode, LCID_JAPANESE)
    CanRepresentInJapaneseCodePage = (StrComp(roundtrip, ch, vbBinaryCompare) = 0)
    Exit Function
Failure:
    CanRepresentInJapaneseCodePage = False
    Err.Clear
End Function

Private Function BuildUnsupportedCharNote() As String
    BuildUnsupportedCharNote = "環境依存文字が使用されているデータがあったので「" & UNSUPPORTED_CHAR_REPLACEMENT & "」に置換しました"
End Function

Private Sub LogErrorsWithStage(ByVal logPrefix As String, ByVal stage As String, ByVal errors As Collection)
    If Not LOG_ENABLED Then Exit Sub
    If errors Is Nothing Then Exit Sub
    If errors.Count = 0 Then Exit Sub
    LogErrorMessage logPrefix & stage & ": " & JoinCollection(errors, " | ")
End Sub

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
    Dim rowSummary As String
    Dim processStage As String
    Dim fatalRowIndex As Long

    processStage = "Initialize"
    fatalRowIndex = 0
    hasUnsupportedCharReplacement = False

    If LOG_ENABLED Then
        logInitialized = False
        logFilePath = vbNullString
        LogInfo "ProcessInstructionRows start. Version=" & VER
        If LenB(logFilePath) > 0 Then
            LogDebug "Log file path=" & logFilePath
        End If
    End If

    processStage = "GetWorksheets"
    Set mainWs = ThisWorkbook.Worksheets("main")
    Set typeWs = ThisWorkbook.Worksheets("type")
    processStage = "BuildTypeCatalog"
    Set typeCatalog = BuildTypeCatalog(typeWs)

    If LOG_ENABLED Then
        If typeCatalog Is Nothing Then
            LogErrorMessage "BuildTypeCatalog returned Nothing"
        Else
            LogDebug "Type catalog count=" & CStr(typeCatalog.Count)
        End If
    End If

    If typeCatalog Is Nothing Or typeCatalog.Count = 0 Then
        MsgBox "typeシートの型定義が取得できません。", vbCritical
        Exit Sub
    End If

    processStage = "PrepareApplication"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    currentRow = 11
    Do While LenB(Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_NO).value))) > 0
        processStage = "RowLoop"
        fatalRowIndex = currentRow
        targetMark = Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_DEF_TARGET).value))
        mainWs.Cells(currentRow, MAIN_COL_MESSAGE).value = vbNullString
        If LOG_ENABLED Then
            rowSummary = "row=" & CStr(currentRow) & _
                          ", defTarget='" & targetMark & "'" & _
                          ", defFile='" & Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_DEF_FILE).value)) & "'" & _
                          ", dataTarget='" & Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_DATA_TARGET).value)) & "'" & _
                          ", dataFile='" & Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_DATA_FILE).value)) & "'"
            LogDebug rowSummary
        End If
        If targetMark = "○" Then
            processStage = "ProcessRow"
            If LOG_ENABLED Then
                LogDebug "Row " & CStr(currentRow) & " is marked for processing"
            End If
            ProcessSingleInstruction mainWs, typeCatalog, currentRow
        Else
            processStage = "SkipRow"
            If LOG_ENABLED Then
                LogDebug "Row " & CStr(currentRow) & " skipped (def target mark='" & targetMark & "')"
            End If
        End If
        currentRow = currentRow + 1
    Loop

    processStage = "LoopCompleted"
    If LOG_ENABLED Then
        LogInfo "ProcessInstructionRows completed. Final row index=" & CStr(currentRow)
    End If

CleanExit:
    processStage = "CleanExit"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

FatalError:
    If LOG_ENABLED Then
        Dim fatalMessage As String
        fatalMessage = "ProcessInstructionRows fatal error(Stage=" & processStage & ", row=" & CStr(fatalRowIndex) & "): ErrNumber=" & CStr(Err.Number) & ", Description=" & Err.Description & ", Source=" & Err.Source
        If LenB(rowSummary) > 0 Then
            fatalMessage = fatalMessage & ", RowSummary=" & rowSummary
        End If
        LogErrorMessage fatalMessage
    End If
    MsgBox "SQL作成中に致命的なエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub ProcessSingleInstruction(ByVal mainWs As Worksheet, ByVal typeCatalog As Object, ByVal rowIndex As Long)
    Dim errors As Collection
    Dim definitionFilePath As String
    Dim definitionSheetName As String
    Dim definitionTableName As String
    Dim colNameAddr As String
    Dim typeAddr As String
    Dim precisionAddr As String
    Dim scaleAddr As String
    Dim pkAddr As String
    Dim notNullAddr As String
    Dim dataFilePath As String
    Dim dataSheetName As String
    Dim dataTableName As String
    Dim dataStartAddr As String
    Dim dbms As String
    Dim dtoLanguage As String
    Dim dtoClassName As String
    Dim definitionWb As Workbook
    Dim definitionWs As Worksheet
    Dim dataWb As Workbook
    Dim dataWs As Worksheet
    Dim columns As Collection
    Dim dataRecords As Collection
    Dim createSql As String
    Dim insertSql As String
    Dim dtoText As String
    Dim createOutputPath As String
    Dim insertOutputPath As String
    Dim dtoOutputPath As String
    Dim timestampText As String
    Dim dataTargetMark As String
    Dim hasDataFile As Boolean
    Dim shouldOutputDto As Boolean
    Dim logPrefix As String
    Dim currentStage As String

    Set errors = New Collection
    hasUnsupportedCharReplacement = False
    On Error GoTo UnexpectedError
    currentStage = "Initialize"

    definitionFilePath = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DEF_FILE).value))
    definitionSheetName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DEF_SHEET).value))
    definitionTableName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DEF_TABLE).value))
    colNameAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_COLNAME_ADDR).value))
    typeAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_TYPE_ADDR).value))
    precisionAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_PRECISION_ADDR).value))
    scaleAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_SCALE_ADDR).value))
    pkAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_PK_ADDR).value))
    notNullAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_NOTNULL_ADDR).value))
    dbms = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DBMS).value))
    dtoLanguage = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DTO_LANG).value))
    dtoClassName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DTO_CLASS).value))
    dataTargetMark = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_TARGET).value))
    dataFilePath = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_FILE).value))
    dataSheetName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_SHEET).value))
    dataTableName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_TABLE).value))
    dataStartAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_START).value))

    logPrefix = "Row " & CStr(rowIndex) & ": "

    If LOG_ENABLED Then
        LogInfo logPrefix & "開始: 定義ファイル=" & definitionFilePath & _
                ", 定義シート=" & definitionSheetName & _
                ", 定義テーブル=" & definitionTableName & _
                ", DBMS=" & dbms & _
                ", データ対象=" & dataTargetMark & _
                ", データファイル=" & dataFilePath & _
                ", データシート=" & dataSheetName & _
                ", データテーブル=" & dataTableName & _
                ", データ開始セル=" & dataStartAddr
    End If

    If LenB(definitionFilePath) = 0 Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "定義ファイルパスが空のためスキップ"
        End If
        Exit Sub
    End If

    hasDataFile = (dataTargetMark = "○")
    shouldOutputDto = (LenB(dtoLanguage) > 0)

    If LOG_ENABLED Then
        LogDebug logPrefix & "hasDataFile=" & CStr(hasDataFile) & ", hasDto=" & CStr(shouldOutputDto)
    End If

    currentStage = "Validation"
    ValidateRequiredValue definitionSheetName, "定義シート名", errors
    ValidateRequiredValue definitionTableName, "定義テーブル名", errors
    ValidateRequiredValue colNameAddr, "カラム名開始セル", errors
    ValidateRequiredValue typeAddr, "型カラム開始セル", errors
    ValidateRequiredValue precisionAddr, "整数桁カラム開始セル", errors
    ValidateRequiredValue scaleAddr, "小数桁カラム開始セル", errors
    ValidateRequiredValue pkAddr, "PKカラム開始セル", errors
    ValidateRequiredValue notNullAddr, "NotNullカラム開始セル", errors
    ValidateDbms dbms, errors

    If shouldOutputDto Then
        ValidateDtoLanguage dtoLanguage, errors
        ValidateRequiredValue dtoClassName, "DTOクラス名", errors
    End If

    If hasDataFile Then
        ValidateRequiredValue dataFilePath, "データファイルパス", errors
        ValidateRequiredValue dataSheetName, "データシート名", errors
        ValidateRequiredValue dataTableName, "データテーブル名", errors
        ValidateRequiredValue dataStartAddr, "データ開始セル", errors
    End If

    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "必須チェック", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    currentStage = "CheckDefinitionFile"
    If Dir$(definitionFilePath, vbNormal) = vbNullString Then
        AddError errors, "定義ファイルパスが存在しません: " & definitionFilePath
        LogErrorsWithStage logPrefix, "定義ファイル存在確認", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    currentStage = "OpenDefinitionWorkbook"
    On Error Resume Next
    Set definitionWb = Application.Workbooks.Open(fileName:=definitionFilePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        AddError errors, "定義ファイルを開けません: " & definitionFilePath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        LogErrorsWithStage logPrefix, "定義ファイルオープン", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    On Error GoTo 0

    If LOG_ENABLED Then
        LogDebug logPrefix & "定義ファイルを開きました: " & definitionFilePath
    End If

    currentStage = "GetDefinitionWorksheet"
    On Error Resume Next
    Set definitionWs = definitionWb.Worksheets(definitionSheetName)
    On Error GoTo 0
    If definitionWs Is Nothing Then
        AddError errors, "定義シートが見つかりません: " & definitionSheetName
        SafeCloseWorkbook definitionWb
        LogErrorsWithStage logPrefix, "定義シート取得", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "定義シートを取得しました: " & definitionSheetName
    End If

    currentStage = "ReadColumnDefinitions"
    Set columns = ReadColumnDefinitions(definitionWs, colNameAddr, typeAddr, precisionAddr, scaleAddr, pkAddr, notNullAddr, typeCatalog, errors)

    SafeCloseWorkbook definitionWb

    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "定義列読み込み", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "有効なカラム定義が1件も取得できません。"
        LogErrorsWithStage logPrefix, "定義列検証", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "定義列数=" & CStr(columns.Count)
    End If

    If hasDataFile Then
        currentStage = "CheckDataFile"
        If Dir$(dataFilePath, vbNormal) = vbNullString Then
            AddError errors, "データファイルパスが存在しません: " & dataFilePath
            LogErrorsWithStage logPrefix, "データファイル存在確認", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            LogDebug logPrefix & "データファイル存在確認OK: " & dataFilePath
        End If

        currentStage = "OpenDataWorkbook"
        On Error Resume Next
        Set dataWb = Application.Workbooks.Open(fileName:=dataFilePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
        If Err.Number <> 0 Then
            AddError errors, "データファイルを開けません: " & dataFilePath & " (" & Err.Description & ")"
            Err.Clear
            On Error GoTo 0
            LogErrorsWithStage logPrefix, "データファイルオープン", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        On Error GoTo 0

        If LOG_ENABLED Then
            LogDebug logPrefix & "データファイルを開きました: " & dataFilePath
        End If

        currentStage = "GetDataWorksheet"
        On Error Resume Next
        Set dataWs = dataWb.Worksheets(dataSheetName)
        On Error GoTo 0
        If dataWs Is Nothing Then
            AddError errors, "データシートが見つかりません: " & dataSheetName
            SafeCloseWorkbook dataWb
            LogErrorsWithStage logPrefix, "データシート取得", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            LogDebug logPrefix & "データシートを取得しました: " & dataSheetName
        End If

        currentStage = "ReadDataRecords"
        Set dataRecords = ReadDataRecords(dataWs, dataStartAddr, columns, errors)

        SafeCloseWorkbook dataWb

        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "投入データ読み込み", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            If dataRecords Is Nothing Then
                LogDebug logPrefix & "投入データ件数=0 (Nothing)"
            Else
                LogDebug logPrefix & "投入データ件数=" & CStr(dataRecords.Count)
            End If
        End If
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQL生成開始"
    End If
    currentStage = "GenerateCreateSql"
    createSql = GenerateCreateSqlText(definitionTableName, columns, dbms, errors)
    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "CREATE SQL生成", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQL生成完了。文字数=" & CStr(Len(createSql))
    End If

    If hasDataFile Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQL生成開始"
        End If
        currentStage = "GenerateInsertSql"
        insertSql = GenerateInsertSqlText(dataTableName, columns, dataRecords, dbms, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "INSERT SQL生成", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQL生成完了。文字数=" & CStr(Len(insertSql))
        End If
    End If

    If shouldOutputDto Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "DTO生成開始"
        End If
        currentStage = "GenerateDtoText"
        dtoText = GenerateDtoText(dtoLanguage, dtoClassName, columns, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "DTO生成", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogDebug logPrefix & "DTO生成完了。文字数=" & CStr(Len(dtoText))
        End If
    End If

    timestampText = Format$(Now, "yyyymmdd_hhnnss")
    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQLファイル出力開始"
    End If
    currentStage = "WriteCreateSqlFile"
    createOutputPath = WriteSqlFile(definitionFilePath, SanitizeForFile(tableName:=definitionTableName) & "_CREATE_" & timestampText & ".sql", createSql, errors)
    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "CREATE SQLファイル出力", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    If LOG_ENABLED Then
        LogInfo logPrefix & "CREATE SQL出力完了: " & createOutputPath
    End If

    If hasDataFile Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQLファイル出力開始"
        End If
        currentStage = "WriteInsertSqlFile"
        insertOutputPath = WriteSqlFile(definitionFilePath, SanitizeForFile(tableName:=dataTableName) & "_INSERT_" & timestampText & ".sql", insertSql, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "INSERT SQLファイル出力", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogInfo logPrefix & "INSERT SQL出力完了: " & insertOutputPath
        End If
    End If

    If shouldOutputDto Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "DTOファイル出力開始"
        End If
        currentStage = "WriteDtoFile"
        dtoOutputPath = WriteDtoFile(definitionFilePath, SanitizeForFile(dtoClassName) & GetDtoFileExtension(dtoLanguage), dtoText, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "DTOファイル出力", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogInfo logPrefix & "DTO出力完了: " & dtoOutputPath
        End If
    End If

    Dim successMessage As String
    If hasDataFile Then
        successMessage = "SQL出力: " & createOutputPath & vbLf & insertOutputPath
    Else
        successMessage = "SQL出力: " & createOutputPath
    End If
    If shouldOutputDto Then
        If LenB(successMessage) > 0 Then
            successMessage = successMessage & vbLf
        End If
        successMessage = successMessage & "DTO出力: " & dtoOutputPath
    End If
    If hasUnsupportedCharReplacement Then
        successMessage = successMessage & vbLf & BuildUnsupportedCharNote()
    End If
    mainWs.Cells(rowIndex, MAIN_COL_MESSAGE).value = successMessage

    If LOG_ENABLED Then
        If hasDataFile And shouldOutputDto Then
            LogInfo logPrefix & "処理完了 (CREATE/INSERT/DTO)"
        ElseIf hasDataFile Then
            LogInfo logPrefix & "処理完了 (CREATE/INSERT)"
        ElseIf shouldOutputDto Then
            LogInfo logPrefix & "処理完了 (CREATE/DTO)"
        Else
            LogInfo logPrefix & "処理完了 (CREATEのみ)"
        End If
    End If

    currentStage = "Finalize"
    hasUnsupportedCharReplacement = False
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    On Error GoTo 0
    If LOG_ENABLED Then
        LogErrorMessage logPrefix & "予期せぬエラー発生(Stage=" & currentStage & "): ErrNumber=" & CStr(Err.Number) & ", Description=" & Err.Description
    End If
    SafeCloseWorkbook definitionWb
    SafeCloseWorkbook dataWb
    hasUnsupportedCharReplacement = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function NormalizeDbms(ByVal dbms As String) As String
    Select Case UCase$(dbms)
        Case "ORACLE"
            NormalizeDbms = "Oracle"
        Case "SQLSERVER"
            NormalizeDbms = "SQLServer"
        Case "SQLITE"
            NormalizeDbms = "SQLite"
        Case "ACCESS"
            NormalizeDbms = "Access"
        Case "H2"
            NormalizeDbms = "H2"
    End Select
End Function

Private Function GenerateInsertSqlText(ByVal tableName As String, ByVal columns As Collection, _
                                       ByVal dataRecords As Collection, ByVal dbms As String, _
                                       ByVal errors As Collection) As String
    Dim schemaName As String
    Dim pureTableName As String
    Dim qualifiedName As String
    Dim columnParts() As String
    Dim idx As Long

    If LOG_ENABLED Then
        Dim columnCountText As String
        Dim recordCountText As String
        If columns Is Nothing Then
            columnCountText = "0"
        Else
            columnCountText = CStr(columns.Count)
        End If
        If dataRecords Is Nothing Then
            recordCountText = "0"
        Else
            recordCountText = CStr(dataRecords.Count)
        End If
        LogDebug "GenerateInsertSqlText開始: table=" & tableName & ", columns=" & columnCountText & ", records=" & recordCountText & ", dbms=" & dbms
    End If

    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "投入データ用のカラム定義が存在しません。"
        Exit Function
    End If

    schemaName = ExtractSchemaName(tableName)
    pureTableName = ExtractTableName(tableName)

    Select Case dbms
        Case "SQLServer"
            If LenB(schemaName) = 0 Then schemaName = "dbo"
        Case "SQLite", "Access"
            schemaName = vbNullString
    End Select

    qualifiedName = BuildQualifiedName(schemaName, pureTableName, dbms)

    ReDim columnParts(0 To columns.Count - 1)
    For idx = 1 To columns.Count
        columnParts(idx - 1) = QuoteIdentifier(columns(idx).Name, dbms)
    Next idx

    Dim columnList As String
    columnList = "(" & Join(columnParts, ", ") & ")"

    If dataRecords Is Nothing Or dataRecords.Count = 0 Then
        If LOG_ENABLED Then
            LogInfo "GenerateInsertSqlText: データ行が存在しないためINSERT文を出力しません"
        End If
        GenerateInsertSqlText = "-- データ行が存在しません。"
        Exit Function
    End If

    Dim statements() As String
    ReDim statements(0 To dataRecords.Count)

    statements(0) = "DELETE FROM " & qualifiedName & ";"

    Dim statementIndex As Long
    statementIndex = 1

    For idx = 1 To dataRecords.Count
        Dim record As Object
        Set record = dataRecords(idx)
        Dim values() As Variant
        Dim addresses() As String
        values = record("Values")
        addresses = record("Addresses")

        If LOG_ENABLED Then
            LogDebug "GenerateInsertSqlText: record=" & CStr(idx) & ", addresses=[" & DescribeStringArray(addresses) & "]"
            LogDebug "GenerateInsertSqlText: record=" & CStr(idx) & ", rawValues=[" & DescribeVariantArray(values) & "]"
        End If

        Dim valueParts() As String
        ReDim valueParts(0 To columns.Count - 1)

        Dim colIdx As Long
        For colIdx = 1 To columns.Count
            Dim def As ColumnDefinition
            Set def = columns(colIdx)
            If LOG_ENABLED Then
                LogDebug "GenerateInsertSqlText: record=" & CStr(idx) & ", column=" & def.Name & ", sourceValue=" & DescribeVariant(values(colIdx)) & ", address=" & addresses(colIdx)
            End If
            valueParts(colIdx - 1) = FormatValueLiteral(def, values(colIdx), addresses(colIdx), dbms, errors)
            If errors.Count > 0 Then
                If LOG_ENABLED Then
                    LogErrorMessage "GenerateInsertSqlText: 値変換エラー record=" & CStr(idx) & ", column=" & def.Name & ", address=" & addresses(colIdx)
                End If
                Exit Function
            End If
            If LOG_ENABLED Then
                LogDebug "GenerateInsertSqlText: record=" & CStr(idx) & ", column=" & def.Name & ", literal=" & valueParts(colIdx - 1)
            End If
        Next colIdx

        statements(statementIndex) = "INSERT INTO " & qualifiedName & " " & columnList & " VALUES (" & Join(valueParts, ", ") & ");"
        statementIndex = statementIndex + 1
    Next idx

    GenerateInsertSqlText = Join(statements, vbCrLf)
End Function

Private Function GenerateDtoText(ByVal dtoLanguage As String, ByVal className As String, ByVal columns As Collection, ByVal errors As Collection) As String
    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "DTO生成対象のカラム定義が存在しません。"
        Exit Function
    End If

    Select Case dtoLanguage
        Case "Java"
            GenerateDtoText = GenerateJavaDtoText(className, columns, errors)
        Case Else
            AddError errors, "DTO出力言語が未対応です: " & dtoLanguage
    End Select
End Function

Private Function GenerateJavaDtoText(ByVal className As String, ByVal columns As Collection, ByVal errors As Collection) As String
    If LenB(className) = 0 Then
        AddError errors, "DTOクラス名が空です。"
        Exit Function
    End If

    Dim imports As Object
    Set imports = CreateObject("Scripting.Dictionary")

    Dim propertyNames() As String
    Dim propertyTypes() As String
    Dim idx As Long

    ReDim propertyNames(1 To columns.Count)
    ReDim propertyTypes(1 To columns.Count)

    For idx = 1 To columns.Count
        Dim def As ColumnDefinition
        Set def = columns(idx)
        propertyNames(idx) = BuildJavaPropertyName(def.Name)
        propertyTypes(idx) = ResolveJavaFieldType(def, imports, errors)
        If errors.Count > 0 Then Exit Function
    Next idx

    Dim lines As New Collection

    If imports.Count > 0 Then
        If imports.Exists("java.math.BigDecimal") Then lines.Add "import java.math.BigDecimal;"
        If imports.Exists("java.time.LocalDate") Then lines.Add "import java.time.LocalDate;"
        If imports.Exists("java.time.LocalTime") Then lines.Add "import java.time.LocalTime;"
        If imports.Exists("java.time.LocalDateTime") Then lines.Add "import java.time.LocalDateTime;"
        lines.Add vbNullString
    End If

    lines.Add "public class " & className & " {"
    lines.Add vbNullString

    For idx = 1 To columns.Count
        lines.Add "    private " & propertyTypes(idx) & " " & propertyNames(idx) & ";"
    Next idx

    lines.Add vbNullString
    lines.Add "    public " & className & "() {"
    lines.Add "    }"

    lines.Add vbNullString

    Dim paramParts() As String
    ReDim paramParts(1 To columns.Count)
    For idx = 1 To columns.Count
        paramParts(idx) = propertyTypes(idx) & " " & propertyNames(idx)
    Next idx
    lines.Add "    public " & className & "(" & Join(paramParts, ", ") & ") {"
    For idx = 1 To columns.Count
        lines.Add "        this." & propertyNames(idx) & " = " & propertyNames(idx) & ";"
    Next idx
    lines.Add "    }"

    lines.Add vbNullString

    For idx = 1 To columns.Count
        Dim pascalName As String
        pascalName = UpperCaseFirst(propertyNames(idx))
        lines.Add "    public " & propertyTypes(idx) & " get" & pascalName & "() {"
        lines.Add "        return " & propertyNames(idx) & ";"
        lines.Add "    }"
        lines.Add vbNullString
        lines.Add "    public void set" & pascalName & "(" & propertyTypes(idx) & " " & propertyNames(idx) & ") {"
        lines.Add "        this." & propertyNames(idx) & " = " & propertyNames(idx) & ";"
        lines.Add "    }"
        lines.Add vbNullString
    Next idx

    lines.Add "    @Override"
    lines.Add "    public String toString() {"
    Dim quoteChar As String
    quoteChar = Chr$(34)
    Dim toStringLine As String
    toStringLine = "        return " & quoteChar & className & "{" & quoteChar & " + "
    For idx = 1 To columns.Count
        If idx > 1 Then
            toStringLine = toStringLine & quoteChar & ", " & quoteChar & " + "
        End If
        If propertyTypes(idx) = "String" Then
            toStringLine = toStringLine & quoteChar & propertyNames(idx) & "='" & quoteChar & " + " & propertyNames(idx) & " + " & quoteChar & "'" & quoteChar & " + "
        Else
            toStringLine = toStringLine & quoteChar & propertyNames(idx) & "=" & quoteChar & " + " & propertyNames(idx) & " + "
        End If
    Next idx
    toStringLine = toStringLine & quoteChar & "}" & quoteChar & ";"
    lines.Add toStringLine
    lines.Add "    }"

    lines.Add "}"

    GenerateJavaDtoText = JoinCollection(lines, vbCrLf)
End Function

Private Function BuildJavaPropertyName(ByVal sourceName As String) As String
    Dim identifier As String
    identifier = NormalizeJavaIdentifier(sourceName)
    If LenB(identifier) = 0 Then
        identifier = "field"
    End If

    BuildJavaPropertyName = LCase$(Left$(identifier, 1)) & Mid$(identifier, 2)
End Function

Private Function NormalizeJavaIdentifier(ByVal sourceName As String) As String
    Dim cleaned As String
    Dim idx As Long

    For idx = 1 To Len(sourceName)
        Dim ch As String
        ch = Mid$(sourceName, idx, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            cleaned = cleaned & ch
        Else
            cleaned = cleaned & "_"
        End If
    Next idx

    If LenB(cleaned) > 0 Then
        Dim firstChar As String
        firstChar = Left$(cleaned, 1)
        If firstChar >= "0" And firstChar <= "9" Then
            cleaned = "_" & cleaned
        End If
    End If

    NormalizeJavaIdentifier = cleaned
End Function

Private Function UpperCaseFirst(ByVal textValue As String) As String
    If LenB(textValue) = 0 Then Exit Function
    UpperCaseFirst = UCase$(Left$(textValue, 1)) & Mid$(textValue, 2)
End Function

Private Function ResolveJavaFieldType(ByVal definition As ColumnDefinition, ByVal imports As Object, ByVal errors As Collection) As String
    Dim typeName As String
    typeName = UCase$(Trim$(definition.DataType))

    Select Case definition.category
        Case CATEGORY_STRING
            ResolveJavaFieldType = "String"
        Case CATEGORY_DATE
            imports("java.time.LocalDate") = True
            ResolveJavaFieldType = "LocalDate"
        Case CATEGORY_TIME
            imports("java.time.LocalTime") = True
            ResolveJavaFieldType = "LocalTime"
        Case CATEGORY_TIMESTAMP
            imports("java.time.LocalDateTime") = True
            ResolveJavaFieldType = "LocalDateTime"
        Case CATEGORY_NUMERIC
            If definition.HasScale And definition.ScaleDigits > 0 Then
                imports("java.math.BigDecimal") = True
                ResolveJavaFieldType = "BigDecimal"
            Else
                Select Case typeName
                    Case "BIGINT"
                        ResolveJavaFieldType = "Long"
                    Case "SMALLINT", "INT", "INTEGER", "TINYINT"
                        ResolveJavaFieldType = "Integer"
                    Case "DOUBLE", "FLOAT", "REAL"
                        ResolveJavaFieldType = "Double"
                    Case Else
                        ResolveJavaFieldType = "Long"
                End Select
            End If
        Case Else
            ResolveJavaFieldType = "String"
    End Select

    If LenB(ResolveJavaFieldType) = 0 Then
        AddError errors, "Javaの型へ変換できません: " & definition.DataType
    End If
End Function

Private Function FormatValueLiteral(ByVal definition As ColumnDefinition, ByVal value As Variant, _
                                    ByVal address As String, ByVal dbms As String, _
                                    ByVal errors As Collection) As String
    If ShouldTreatAsNull(definition, value) Then
        FormatValueLiteral = "NULL"
        Exit Function
    End If

    Select Case definition.category
        Case CATEGORY_NUMERIC
            FormatValueLiteral = FormatNumericLiteral(value)
        Case CATEGORY_STRING
            FormatValueLiteral = FormatStringLiteral(value)
        Case CATEGORY_DATE
            FormatValueLiteral = FormatDateLiteral(value, address, dbms, errors)
        Case CATEGORY_TIME
            FormatValueLiteral = FormatTimeLiteral(value, address, dbms, errors)
        Case CATEGORY_TIMESTAMP
            FormatValueLiteral = FormatTimestampLiteral(value, address, dbms, errors)
        Case Else
            FormatValueLiteral = FormatStringLiteral(value)
    End Select
End Function

Private Function ShouldTreatAsNull(ByVal definition As ColumnDefinition, ByVal value As Variant) As Boolean
    If IsNull(value) Or IsEmpty(value) Then
        ShouldTreatAsNull = True
        Exit Function
    End If

    If VarType(value) = vbString Then
        Dim textValue As String
        textValue = CStr(value)

        Select Case definition.category
            Case CATEGORY_STRING
                ShouldTreatAsNull = (LenB(textValue) = 0)
            Case Else
                ShouldTreatAsNull = (LenB(Trim$(textValue)) = 0)
        End Select
    Else
        ShouldTreatAsNull = False
    End If
End Function
Private Function HasValueWithoutTrim(ByVal value As Variant) As Boolean
    If IsNull(value) Or IsEmpty(value) Then
        HasValueWithoutTrim = False
    ElseIf VarType(value) = vbString Then
        HasValueWithoutTrim = (LenB(CStr(value)) > 0)
    Else
        HasValueWithoutTrim = True
    End If
End Function


Private Function IsNullOrEmptyValue(ByVal value As Variant) As Boolean
    If IsNull(value) Or IsEmpty(value) Then
        IsNullOrEmptyValue = True
    ElseIf VarType(value) = vbString Then
        IsNullOrEmptyValue = (LenB(Trim$(CStr(value))) = 0)
    Else
        IsNullOrEmptyValue = False
    End If
End Function

Private Function FormatNumericLiteral(ByVal value As Variant) As String
    FormatNumericLiteral = Trim$(CStr(value))
End Function

Private Function FormatStringLiteral(ByVal value As Variant) As String
    Dim textValue As String
    textValue = ReplaceUnsupportedCharacters(CStr(value))
    textValue = Replace(textValue, "'", "''")
    FormatStringLiteral = "'" & textValue & "'"
End Function

Private Function FormatDateLiteral(ByVal value As Variant, ByVal address As String, _
                                   ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "日付の値を解析できません: " & address
        Exit Function
    End If

    Dim dateText As String
    dateText = Format$(dt, "yyyy-mm-dd")

    Select Case dbms
        Case "Oracle"
            FormatDateLiteral = "TO_DATE('" & dateText & "','YYYY-MM-DD')"
        Case "SQLServer"
            FormatDateLiteral = "CAST('" & dateText & "' AS DATETIME)"
        Case "Access"
            FormatDateLiteral = "#" & dateText & "#"
        Case "H2"
            FormatDateLiteral = "DATE '" & dateText & "'"
        Case Else
            FormatDateLiteral = "'" & dateText & "'"
    End Select
End Function

Private Function FormatTimeLiteral(ByVal value As Variant, ByVal address As String, _
                                   ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "時刻の値を解析できません: " & address
        Exit Function
    End If

    Dim timeText As String
    timeText = Format$(dt, "HH:nn:ss")

    Select Case dbms
        Case "Oracle"
            FormatTimeLiteral = "TO_TIMESTAMP('" & ORACLE_TIME_BASE_DATE & " " & timeText & "','YYYY-MM-DD HH24:MI:SS')"
        Case "SQLServer"
            FormatTimeLiteral = "CAST('1900-01-01 " & timeText & "' AS DATETIME)"
        Case "Access"
            FormatTimeLiteral = "#" & timeText & "#"
        Case "H2"
            FormatTimeLiteral = "TIME '" & timeText & "'"
        Case Else
            FormatTimeLiteral = "'" & timeText & "'"
    End Select
End Function


Private Function ResolveDbmsTypeName(ByVal rawTypeName As String, ByVal category As String, ByVal dbms As String, ByVal errors As Collection) As String
    Dim normalized As String
    normalized = UCase$(Trim$(rawTypeName))
    Dim resolved As String

    Select Case dbms
        Case "SQLServer"
            resolved = ResolveSqlServerTypeName(normalized, category)
        Case "Oracle"
            resolved = ResolveOracleTypeName(normalized, category)
        Case "SQLite"
            resolved = ResolveSqliteTypeName(normalized, category)
        Case "Access"
            resolved = ResolveAccessTypeName(normalized, category)
        Case "H2"
            resolved = ResolveH2TypeName(normalized, category)
        Case Else
            resolved = normalized
    End Select

    If LenB(resolved) = 0 Then
        AddError errors, "DBMS(" & dbms & ")で利用できない型です: " & rawTypeName
    End If

    ResolveDbmsTypeName = resolved
End Function

Private Function ResolveSqlServerTypeName(ByVal normalized As String, ByVal category As String) As String
    Select Case normalized
        Case "NUMBER", "DECIMAL"
            ResolveSqlServerTypeName = "DECIMAL"
        Case "VARCHAR2", "NVARCHAR"
            ResolveSqlServerTypeName = "NVARCHAR"
        Case "VARCHAR"
            ResolveSqlServerTypeName = "VARCHAR"
        Case "CHAR"
            ResolveSqlServerTypeName = "CHAR"
        Case "TIMESTAMP"
            ResolveSqlServerTypeName = "DATETIME"
        Case "DATE"
            ResolveSqlServerTypeName = "DATETIME"
        Case "TIME"
            ResolveSqlServerTypeName = "DATETIME"
        Case "DATETIME", "DATETIME2"
            ResolveSqlServerTypeName = normalized
        Case Else
            Select Case category
                Case CATEGORY_NUMERIC
                    ResolveSqlServerTypeName = "DECIMAL"
                Case CATEGORY_STRING
                    ResolveSqlServerTypeName = "NVARCHAR"
                Case CATEGORY_DATE
                    ResolveSqlServerTypeName = "DATETIME"
                Case CATEGORY_TIME
                    ResolveSqlServerTypeName = "DATETIME"
                Case CATEGORY_TIMESTAMP
                    ResolveSqlServerTypeName = "DATETIME"
                Case Else
                    ResolveSqlServerTypeName = normalized
            End Select
    End Select
End Function

Private Function ResolveOracleTypeName(ByVal normalized As String, ByVal category As String) As String
    Select Case normalized
        Case "DECIMAL"
            ResolveOracleTypeName = "NUMBER"
        Case "NUMBER"
            ResolveOracleTypeName = "NUMBER"
        Case "VARCHAR"
            ResolveOracleTypeName = "VARCHAR2"
        Case "NVARCHAR"
            ResolveOracleTypeName = "NVARCHAR2"
        Case "VARCHAR2"
            ResolveOracleTypeName = "VARCHAR2"
        Case "CHAR"
            ResolveOracleTypeName = "CHAR"
        Case "DATE"
            ResolveOracleTypeName = "DATE"
        Case "TIME"
            ResolveOracleTypeName = "DATE"
        Case "TIMESTAMP"
            ResolveOracleTypeName = "TIMESTAMP"
        Case Else
            Select Case category
                Case CATEGORY_NUMERIC
                    ResolveOracleTypeName = "NUMBER"
                Case CATEGORY_STRING
                    ResolveOracleTypeName = "VARCHAR2"
                Case CATEGORY_DATE
                    ResolveOracleTypeName = "DATE"
                Case CATEGORY_TIME
                    ResolveOracleTypeName = "DATE"
                Case CATEGORY_TIMESTAMP
                    ResolveOracleTypeName = "TIMESTAMP"
                Case Else
                    ResolveOracleTypeName = normalized
            End Select
    End Select
End Function

Private Function ResolveSqliteTypeName(ByVal normalized As String, ByVal category As String) As String
    Select Case normalized
        Case "NUMBER", "DECIMAL"
            ResolveSqliteTypeName = "NUMERIC"
        Case "VARCHAR", "VARCHAR2", "NVARCHAR", "CHAR"
            ResolveSqliteTypeName = "TEXT"
        Case "DATE", "TIME", "TIMESTAMP"
            ResolveSqliteTypeName = "TEXT"
        Case Else
            Select Case category
                Case CATEGORY_NUMERIC
                    ResolveSqliteTypeName = "NUMERIC"
                Case CATEGORY_STRING
                    ResolveSqliteTypeName = "TEXT"
                Case CATEGORY_DATE, CATEGORY_TIME, CATEGORY_TIMESTAMP
                    ResolveSqliteTypeName = "TEXT"
                Case Else
                    ResolveSqliteTypeName = normalized
            End Select
    End Select
End Function

Private Function ResolveAccessTypeName(ByVal normalized As String, ByVal category As String) As String
    Select Case normalized
        Case "NUMBER", "DECIMAL"
            ResolveAccessTypeName = "DECIMAL"
        Case "INTEGER", "INT"
            ResolveAccessTypeName = "INTEGER"
        Case "VARCHAR", "VARCHAR2", "NVARCHAR", "CHAR"
            ResolveAccessTypeName = "TEXT"
        Case "TIMESTAMP", "DATE", "TIME", "DATETIME"
            ResolveAccessTypeName = "DATETIME"
        Case Else
            Select Case category
                Case CATEGORY_NUMERIC
                    ResolveAccessTypeName = "DECIMAL"
                Case CATEGORY_STRING
                    ResolveAccessTypeName = "TEXT"
                Case CATEGORY_DATE, CATEGORY_TIME, CATEGORY_TIMESTAMP
                    ResolveAccessTypeName = "DATETIME"
                Case Else
                    ResolveAccessTypeName = normalized
            End Select
    End Select
End Function

Private Function ResolveH2TypeName(ByVal normalized As String, ByVal category As String) As String
    Select Case normalized
        Case "NUMBER", "DECIMAL"
            ResolveH2TypeName = "DECIMAL"
        Case "INTEGER", "INT"
            ResolveH2TypeName = "INTEGER"
        Case "VARCHAR", "VARCHAR2", "NVARCHAR"
            ResolveH2TypeName = "VARCHAR"
        Case "CHAR"
            ResolveH2TypeName = "CHAR"
        Case "DATE"
            ResolveH2TypeName = "DATE"
        Case "TIME"
            ResolveH2TypeName = "TIME"
        Case "TIMESTAMP", "DATETIME"
            ResolveH2TypeName = "TIMESTAMP"
        Case Else
            Select Case category
                Case CATEGORY_NUMERIC
                    ResolveH2TypeName = "DECIMAL"
                Case CATEGORY_STRING
                    ResolveH2TypeName = "VARCHAR"
                Case CATEGORY_DATE
                    ResolveH2TypeName = "DATE"
                Case CATEGORY_TIME
                    ResolveH2TypeName = "TIME"
                Case CATEGORY_TIMESTAMP
                    ResolveH2TypeName = "TIMESTAMP"
                Case Else
                    ResolveH2TypeName = normalized
            End Select
    End Select
End Function

Private Function FormatTimestampLiteral(ByVal value As Variant, ByVal address As String, _
                                        ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "日時の値を解析できません: " & address
        Exit Function
    End If

    Dim timestampText As String
    timestampText = Format$(dt, "yyyy-mm-dd HH:nn:ss")

    Select Case dbms
        Case "Oracle"
            FormatTimestampLiteral = "TO_TIMESTAMP('" & timestampText & "','YYYY-MM-DD HH24:MI:SS')"
        Case "SQLServer"
            FormatTimestampLiteral = "CAST('" & timestampText & "' AS DATETIME)"
        Case "Access"
            FormatTimestampLiteral = "#" & timestampText & "#"
        Case "H2"
            FormatTimestampLiteral = "TIMESTAMP '" & timestampText & "'"
        Case Else
            FormatTimestampLiteral = "'" & timestampText & "'"
    End Select
End Function

Private Function TryParseDateValue(ByVal value As Variant, ByRef resultValue As Date) As Boolean
    On Error Resume Next
    If IsDate(value) Then
        resultValue = CDate(value)
        TryParseDateValue = True
    Else
        Dim textValue As String
        textValue = Trim$(CStr(value))
        If LenB(textValue) = 0 Then
            TryParseDateValue = False
        Else
            resultValue = CDate(textValue)
            If Err.Number = 0 Then
                TryParseDateValue = True
            End If
        End If
    End If

    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo 0
End Function
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
        columnName = Trim$(CStr(targetWs.Cells(currentRow, nameCell.Column).value))
        If LenB(columnName) = 0 Then Exit Do

        Dim columnType As String
        columnType = Trim$(CStr(targetWs.Cells(currentRow, typeCell.Column).value))
        If LenB(columnType) = 0 Then
            AddError errors, "型が未入力です: " & targetWs.Cells(currentRow, typeCell.Column).address(False, False)
            Exit Do
        End If

        Dim typeKey As String
        typeKey = UCase$(columnType)
        If Not typeCatalog.Exists(typeKey) Then
            AddError errors, "型がtypeシートに存在しません: " & columnType & " (" & targetWs.Cells(currentRow, typeCell.Column).address(False, False) & ")"
            Exit Do
        End If

        Dim category As String
        category = CStr(typeCatalog(typeKey))

        Dim definition As ColumnDefinition
        Set definition = New ColumnDefinition
        definition.Name = columnName
        definition.DataType = columnType
        definition.category = category

        Dim precisionText As String
        precisionText = Trim$(CStr(targetWs.Cells(currentRow, precisionCell.Column).value))
        Dim scaleText As String
        scaleText = Trim$(CStr(targetWs.Cells(currentRow, scaleCell.Column).value))
        Dim isIntegerNumeric As Boolean
        isIntegerNumeric = (category = CATEGORY_NUMERIC And IsIntegerLikeType(typeKey))
        Dim parsedPrecision As Long
        Dim parsedScale As Long

        If category = CATEGORY_STRING Then
            If Not TryParsePositiveInteger(precisionText, parsedPrecision) Then
                AddError errors, "整数桁が不正です: " & targetWs.Cells(currentRow, precisionCell.Column).address(False, False)
            Else
                definition.Precision = parsedPrecision
                definition.HasPrecision = True
            End If
        ElseIf category = CATEGORY_NUMERIC Then
            If LenB(precisionText) > 0 Then
                If Not TryParsePositiveInteger(precisionText, parsedPrecision) Then
                    AddError errors, "整数桁が不正です: " & targetWs.Cells(currentRow, precisionCell.Column).address(False, False)
                Else
                    definition.Precision = parsedPrecision
                    definition.HasPrecision = True
                End If
            ElseIf Not isIntegerNumeric Then
                AddError errors, "整数桁が不正です: " & targetWs.Cells(currentRow, precisionCell.Column).address(False, False)
            End If

            If LenB(scaleText) > 0 Then
                If Not TryParseNonNegativeInteger(scaleText, parsedScale) Then
                    AddError errors, "小数桁が不正です: " & targetWs.Cells(currentRow, scaleCell.Column).address(False, False)
                Else
                    definition.ScaleDigits = parsedScale
                    definition.HasScale = True
                End If
            ElseIf Not isIntegerNumeric Then
                AddError errors, "小数桁が不正です: " & targetWs.Cells(currentRow, scaleCell.Column).address(False, False)
            End If
        End If

        If errors.Count > 0 Then
            Exit Do
        End If

        Dim pkValue As Variant
        pkValue = targetWs.Cells(currentRow, pkCell.Column).value
        definition.IsPrimaryKey = HasValueWithoutTrim(pkValue)

        Dim notNullValue As Variant
        notNullValue = targetWs.Cells(currentRow, notNullCell.Column).value
        definition.IsNotNull = HasValueWithoutTrim(notNullValue) Or definition.IsPrimaryKey

        columns.Add definition
        currentRow = currentRow + 1
    Loop

    If errors.Count = 0 Then
        Set ReadColumnDefinitions = columns
    End If
End Function

Private Function ReadDataRecords(ByVal targetWs As Worksheet, ByVal startAddr As String, _
                                 ByVal columns As Collection, ByVal errors As Collection) As Collection
    Dim startCell As Range
    Dim headerRow As Long
    Dim dataRow As Long
    Dim noColumn As Long
    Dim headerValue As String

    Set startCell = ResolveSingleCell(targetWs, startAddr, "データ開始セル", errors)
    If errors.Count > 0 Then Exit Function
    If startCell Is Nothing Then Exit Function

    noColumn = startCell.Column
    headerValue = Trim$(CStr(startCell.value))

    If StrComp(headerValue, "No", vbTextCompare) = 0 Then
        noColumn = startCell.Column
        headerRow = startCell.Row
        dataRow = headerRow + 1
    Else
        Dim headerCandidateRow As Long
        headerCandidateRow = startCell.Row - 1
        If headerCandidateRow < 1 Then
            AddError errors, "No列ヘッダ位置を特定できません: " & startCell.address(False, False)
            Exit Function
        End If

        noColumn = LocateNoHeaderColumn(targetWs, headerCandidateRow, startCell.Column)
        If noColumn = 0 Then
            AddError errors, "No列ヘッダが見つかりません: " & targetWs.Cells(headerCandidateRow, startCell.Column).address(False, False)
            Exit Function
        End If

        headerRow = headerCandidateRow
        dataRow = startCell.Row
    End If

    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare

    Dim currentColumn As Long
    currentColumn = noColumn
    Do While LenB(Trim$(CStr(targetWs.Cells(headerRow, currentColumn).value))) > 0
        Dim colName As String
        colName = Trim$(CStr(targetWs.Cells(headerRow, currentColumn).value))
        If LenB(colName) > 0 Then
            If Not headerMap.Exists(colName) Then
                headerMap.Add colName, currentColumn
            End If
        End If
        currentColumn = currentColumn + 1
    Loop

    Dim idx As Long
    For idx = 1 To columns.Count
        Dim def As ColumnDefinition
        Set def = columns(idx)
        If Not headerMap.Exists(def.Name) Then
            AddError errors, "投入データ定義ヘッダにカラムが存在しません: " & def.Name
        End If
    Next idx

    If errors.Count > 0 Then Exit Function

    Dim records As New Collection
    Dim rowValues() As Variant
    Dim rowAddresses() As String
    Dim currentRow As Long
    currentRow = dataRow

    Do While currentRow <= targetWs.Rows.Count
        Dim noValue As String
        noValue = Trim$(CStr(targetWs.Cells(currentRow, noColumn).value))
        If LenB(noValue) = 0 Then Exit Do

        ReDim rowValues(1 To columns.Count)
        ReDim rowAddresses(1 To columns.Count)

        For idx = 1 To columns.Count
            Dim valueCell As Range
            Dim keyName As String
            keyName = CStr(columns(idx).Name)
            Set valueCell = targetWs.Cells(currentRow, CLng(headerMap(keyName)))
            Dim cellValue As Variant
            cellValue = valueCell.value
            rowValues(idx) = cellValue
            rowAddresses(idx) = valueCell.address(False, False)

            If columns(idx).IsPrimaryKey Then
                If Not HasValueWithoutTrim(cellValue) Then
                    AddError errors, "必須項目が空欄です: " & valueCell.address(False, False)
                End If
            ElseIf columns(idx).IsNotNull Then
                If Not HasValueWithoutTrim(cellValue) Then
                    AddError errors, "必須項目が空欄です: " & valueCell.address(False, False)
                End If
            End If
        Next idx

        If errors.Count > 0 Then Exit Do

        Dim record As Object
        Set record = CreateObject("Scripting.Dictionary")
        record.CompareMode = vbBinaryCompare
        record.Add "Values", rowValues
        record.Add "Addresses", rowAddresses
        records.Add record

        currentRow = currentRow + 1
    Loop

    If errors.Count = 0 Then
        Set ReadDataRecords = records
    End If
End Function

Private Function LocateNoHeaderColumn(ByVal targetWs As Worksheet, ByVal headerRow As Long, _
                                      ByVal hintColumn As Long) As Long
    Dim valueText As String
    valueText = Trim$(CStr(targetWs.Cells(headerRow, hintColumn).value))
    If StrComp(valueText, "No", vbTextCompare) = 0 Then
        LocateNoHeaderColumn = hintColumn
        Exit Function
    End If

    Dim searchColumn As Long
    For searchColumn = hintColumn - 1 To 1 Step -1
        valueText = Trim$(CStr(targetWs.Cells(headerRow, searchColumn).value))
        If StrComp(valueText, "No", vbTextCompare) = 0 Then
            LocateNoHeaderColumn = searchColumn
            Exit Function
        End If
    Next searchColumn

    Dim lastColumn As Long
    lastColumn = targetWs.Cells(headerRow, targetWs.columns.Count).End(xlToLeft).Column
    If lastColumn < hintColumn Then
        lastColumn = hintColumn
    End If

    For searchColumn = hintColumn + 1 To lastColumn
        valueText = Trim$(CStr(targetWs.Cells(headerRow, searchColumn).value))
        If StrComp(valueText, "No", vbTextCompare) = 0 Then
            LocateNoHeaderColumn = searchColumn
            Exit Function
        End If
    Next searchColumn
End Function

Private Function GenerateCreateSqlText(ByVal tableName As String, ByVal columns As Collection, ByVal dbms As String, ByVal errors As Collection) As String
    Dim schemaName As String
    Dim pureTableName As String
    Dim qualifiedName As String

    schemaName = ExtractSchemaName(tableName)
    pureTableName = ExtractTableName(tableName)

    Select Case dbms
        Case "SQLServer"
            If LenB(schemaName) = 0 Then schemaName = "dbo"
        Case "SQLite", "Access"
            schemaName = vbNullString
    End Select

    qualifiedName = BuildQualifiedName(schemaName, pureTableName, dbms)

    Dim lines() As String
    ReDim lines(0 To columns.Count)
    Dim idx As Long
    Dim pkColumns As New Collection

    For idx = 1 To columns.Count
        Dim def As ColumnDefinition
        Set def = columns(idx)
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

    GenerateCreateSqlText = dropSql & vbCrLf & vbCrLf & createSql
End Function

Private Function IsIntegerLikeType(ByVal typeName As String) As Boolean
    Select Case UCase$(Trim$(typeName))
        Case "INT", "INTEGER", "BIGINT", "SMALLINT", "TINYINT"
            IsIntegerLikeType = True
    End Select
End Function

Private Function FormatDataType(ByVal definition As ColumnDefinition, ByVal dbms As String, ByVal errors As Collection) As String
    Dim baseType As String
    baseType = ResolveDbmsTypeName(definition.DataType, definition.category, dbms, errors)
    If LenB(baseType) = 0 Or errors.Count > 0 Then Exit Function

    Select Case definition.category
        Case CATEGORY_NUMERIC
            If IsIntegerLikeType(definition.DataType) Then
                FormatDataType = baseType
            Else
                If Not definition.HasPrecision Or Not definition.HasScale Then
                    AddError errors, "数値型の桁情報が不足しています: " & definition.Name
                    Exit Function
                End If
                FormatDataType = baseType & "(" & CStr(definition.Precision) & "," & CStr(definition.ScaleDigits) & ")"
            End If
        Case CATEGORY_STRING
            If Not definition.HasPrecision Then
                AddError errors, "文字列型の桁情報が不足しています: " & definition.Name
                Exit Function
            End If
            If dbms = "SQLite" And UCase$(baseType) = "TEXT" Then
                FormatDataType = baseType
            Else
                FormatDataType = baseType & "(" & CStr(definition.Precision) & ")"
            End If
        Case Else
            FormatDataType = baseType
    End Select
End Function

Private Function BuildDropStatement(ByVal qualifiedName As String, ByVal dbms As String) As String
    Select Case dbms
        Case "SQLServer"
            Dim objectName As String
            objectName = Replace(qualifiedName, "[", vbNullString)
            objectName = Replace(objectName, "]", vbNullString)
            BuildDropStatement = "IF OBJECT_ID(N'" & objectName & "', N'U') IS NOT NULL" & vbCrLf & _
                               "    DROP TABLE " & qualifiedName & ";"
        Case "SQLite"
            BuildDropStatement = "DROP TABLE IF EXISTS " & qualifiedName & ";"
        Case "H2"
            BuildDropStatement = "DROP TABLE IF EXISTS " & qualifiedName & ";"
        Case Else
            BuildDropStatement = "DROP TABLE " & qualifiedName & ";"
    End Select
End Function

Private Function WriteSqlFile(ByVal sourceFilePath As String, ByVal fileName As String, _
                              ByVal sqlText As String, ByVal errors As Collection) As String
    WriteSqlFile = WriteTextFile(sourceFilePath, fileName, sqlText, errors, "SQLファイル")
End Function

Private Function WriteDtoFile(ByVal sourceFilePath As String, ByVal fileName As String, _
                              ByVal dtoText As String, ByVal errors As Collection) As String
    WriteDtoFile = WriteTextFile(sourceFilePath, fileName, dtoText, errors, "DTOファイル")
End Function

Private Function WriteTextFile(ByVal sourceFilePath As String, ByVal fileName As String, _
                              ByVal content As String, ByVal errors As Collection, _
                              ByVal label As String) As String
    Dim fso As Object
    Dim parentFolder As String
    Dim outputPath As String
    Dim ts As Object

    If LenB(fileName) = 0 Then
        AddError errors, "出力ファイル名が正しく生成できません。"
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(sourceFilePath)
    If LenB(parentFolder) = 0 Then
        AddError errors, "参照ファイルの親フォルダを取得できません: " & sourceFilePath
        Exit Function
    End If

    outputPath = fso.BuildPath(parentFolder, fileName)

    On Error Resume Next
    Set ts = fso.CreateTextFile(outputPath, True, False)
    If Err.Number <> 0 Then
        AddError errors, label & "を作成できません: " & outputPath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ts.Write content
    ts.Close

    WriteTextFile = outputPath
End Function

Private Function BuildTypeCatalog(ByVal typeWs As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim currentRow As Long
    currentRow = 3

    Do While LenB(Trim$(CStr(typeWs.Cells(currentRow, MAIN_COL_NO).value))) > 0
        Dim category As String
        category = Trim$(CStr(typeWs.Cells(currentRow, MAIN_COL_NO).value))
        Dim currentColumn As Long
        currentColumn = 3
        Do While LenB(Trim$(CStr(typeWs.Cells(currentRow, currentColumn).value))) > 0
            Dim typeName As String
            typeName = UCase$(Trim$(CStr(typeWs.Cells(currentRow, currentColumn).value)))
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

Private Sub ValidateDbms(ByRef dbms As String, ByVal errors As Collection)
    Dim normalized As String
    normalized = NormalizeDbms(dbms)

    If LenB(normalized) = 0 Then
        AddError errors, "DBMSが不正です: " & dbms
    Else
        dbms = normalized
    End If
End Sub

Private Sub ValidateDtoLanguage(ByRef dtoLanguage As String, ByVal errors As Collection)
    Dim normalized As String
    normalized = NormalizeDtoLanguage(dtoLanguage)

    If LenB(normalized) = 0 Then
        AddError errors, "DTO出力言語が不正です: " & dtoLanguage
    Else
        dtoLanguage = normalized
    End If
End Sub

Private Function NormalizeDtoLanguage(ByVal dtoLanguage As String) As String
    Select Case UCase$(dtoLanguage)
        Case "JAVA"
            NormalizeDtoLanguage = "Java"
    End Select
End Function

Private Function GetDtoFileExtension(ByVal dtoLanguage As String) As String
    Select Case dtoLanguage
        Case "Java"
            GetDtoFileExtension = ".java"
    End Select
End Function

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
        Case "SQLServer", "Access"
            cleaned = Replace(cleaned, "[", vbNullString)
            cleaned = Replace(cleaned, "]", vbNullString)
            QuoteIdentifier = "[" & cleaned & "]"
        Case "H2"
            cleaned = Replace(cleaned, """", vbNullString)
            QuoteIdentifier = cleaned
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
    Dim messageText As String
    messageText = JoinCollection(errors, vbLf)
    If hasUnsupportedCharReplacement Then
        If LenB(messageText) > 0 Then
            messageText = messageText & vbLf
        End If
        messageText = messageText & BuildUnsupportedCharNote()
    End If
    mainWs.Cells(rowIndex, MAIN_COL_MESSAGE).value = messageText
    If LOG_ENABLED Then
        Dim flattened As String
        flattened = Replace(messageText, vbLf, " | ")
        LogErrorMessage "Row " & CStr(rowIndex) & " エラー: " & flattened
    End If
    hasUnsupportedCharReplacement = False
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










