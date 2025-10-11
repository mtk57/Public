Attribute VB_Name = "Module_SQLCreator"
Option Explicit

Private Const VER As String = "2.2.6"

Private Const LOG_ENABLED As Boolean = True
Private Const LOG_FILE_NAME As String = "SQLCreator_debug.log"

Private logFilePath As String
Private logInitialized As Boolean

Private Const CATEGORY_NUMERIC As String = "���l"
Private Const CATEGORY_STRING As String = "������"
Private Const CATEGORY_DATE As String = "���t"
Private Const CATEGORY_TIME As String = "����"
Private Const CATEGORY_TIMESTAMP As String = "����"

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
Private Const MAIN_COL_MESSAGE As Long = 20

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

Private Sub LogErrorsWithStage(ByVal logPrefix As String, ByVal stage As String, ByVal errors As Collection)
    If Not LOG_ENABLED Then Exit Sub
    If errors Is Nothing Then Exit Sub
    If errors.Count = 0 Then Exit Sub
    LogErrorMessage logPrefix & stage & ": " & JoinCollection(errors, " | ")
End Sub

Public Sub �{�^��1_Click()
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

    If LOG_ENABLED Then
        logInitialized = False
        logFilePath = vbNullString
        LogInfo "ProcessInstructionRows start. Version=" & VER
        If LenB(logFilePath) > 0 Then
            LogDebug "Log file path=" & logFilePath
        End If
    End If

    Set mainWs = ThisWorkbook.Worksheets("main")
    Set typeWs = ThisWorkbook.Worksheets("type")
    Set typeCatalog = BuildTypeCatalog(typeWs)

    If LOG_ENABLED Then
        If typeCatalog Is Nothing Then
            LogErrorMessage "BuildTypeCatalog returned Nothing"
        Else
            LogDebug "Type catalog count=" & CStr(typeCatalog.Count)
        End If
    End If

    If typeCatalog Is Nothing Or typeCatalog.Count = 0 Then
        MsgBox "type�V�[�g�̌^��`���擾�ł��܂���B", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    currentRow = 11
    Do While LenB(Trim$(CStr(mainWs.Cells(currentRow, MAIN_COL_NO).value))) > 0
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
        If targetMark = "��" Then
            If LOG_ENABLED Then
                LogDebug "Row " & CStr(currentRow) & " is marked for processing"
            End If
            ProcessSingleInstruction mainWs, typeCatalog, currentRow
        Else
            If LOG_ENABLED Then
                LogDebug "Row " & CStr(currentRow) & " skipped (def target mark='" & targetMark & "')"
            End If
        End If
        currentRow = currentRow + 1
    Loop

    If LOG_ENABLED Then
        LogInfo "ProcessInstructionRows completed. Final row index=" & CStr(currentRow)
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

FatalError:
    If LOG_ENABLED Then
        LogErrorMessage "ProcessInstructionRows fatal error: ErrNumber=" & CStr(Err.Number) & ", Description=" & Err.Description
    End If
    MsgBox "SQL�쐬���ɒv���I�ȃG���[���������܂���: " & Err.Description, vbCritical
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
    Dim definitionWb As Workbook
    Dim definitionWs As Worksheet
    Dim dataWb As Workbook
    Dim dataWs As Worksheet
    Dim columns As Collection
    Dim dataRecords As Collection
    Dim createSql As String
    Dim insertSql As String
    Dim createOutputPath As String
    Dim insertOutputPath As String
    Dim timestampText As String
    Dim dataTargetMark As String
    Dim hasDataFile As Boolean
    Dim logPrefix As String
    Dim currentStage As String

    Set errors = New Collection
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
    dataTargetMark = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_TARGET).value))
    dataFilePath = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_FILE).value))
    dataSheetName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_SHEET).value))
    dataTableName = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_TABLE).value))
    dataStartAddr = Trim$(CStr(mainWs.Cells(rowIndex, MAIN_COL_DATA_START).value))

    logPrefix = "Row " & CStr(rowIndex) & ": "

    If LOG_ENABLED Then
        LogInfo logPrefix & "�J�n: ��`�t�@�C��=" & definitionFilePath & _
                ", ��`�V�[�g=" & definitionSheetName & _
                ", ��`�e�[�u��=" & definitionTableName & _
                ", DBMS=" & dbms & _
                ", �f�[�^�Ώ�=" & dataTargetMark & _
                ", �f�[�^�t�@�C��=" & dataFilePath & _
                ", �f�[�^�V�[�g=" & dataSheetName & _
                ", �f�[�^�e�[�u��=" & dataTableName & _
                ", �f�[�^�J�n�Z��=" & dataStartAddr
    End If

    If LenB(definitionFilePath) = 0 Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "��`�t�@�C���p�X����̂��߃X�L�b�v"
        End If
        Exit Sub
    End If

    hasDataFile = (dataTargetMark = "��")

    If LOG_ENABLED Then
        LogDebug logPrefix & "hasDataFile=" & CStr(hasDataFile)
    End If

    currentStage = "Validation"
    ValidateRequiredValue definitionSheetName, "��`�V�[�g��", errors
    ValidateRequiredValue definitionTableName, "��`�e�[�u����", errors
    ValidateRequiredValue colNameAddr, "�J�������J�n�Z��", errors
    ValidateRequiredValue typeAddr, "�^�J�����J�n�Z��", errors
    ValidateRequiredValue precisionAddr, "�������J�����J�n�Z��", errors
    ValidateRequiredValue scaleAddr, "�������J�����J�n�Z��", errors
    ValidateRequiredValue pkAddr, "PK�J�����J�n�Z��", errors
    ValidateRequiredValue notNullAddr, "NotNull�J�����J�n�Z��", errors
    ValidateDbms dbms, errors

    If hasDataFile Then
        ValidateRequiredValue dataFilePath, "�f�[�^�t�@�C���p�X", errors
        ValidateRequiredValue dataSheetName, "�f�[�^�V�[�g��", errors
        ValidateRequiredValue dataTableName, "�f�[�^�e�[�u����", errors
        ValidateRequiredValue dataStartAddr, "�f�[�^�J�n�Z��", errors
    End If

    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "�K�{�`�F�b�N", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    currentStage = "CheckDefinitionFile"
    If Dir$(definitionFilePath, vbNormal) = vbNullString Then
        AddError errors, "��`�t�@�C���p�X�����݂��܂���: " & definitionFilePath
        LogErrorsWithStage logPrefix, "��`�t�@�C�����݊m�F", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    currentStage = "OpenDefinitionWorkbook"
    On Error Resume Next
    Set definitionWb = Application.Workbooks.Open(fileName:=definitionFilePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        AddError errors, "��`�t�@�C�����J���܂���: " & definitionFilePath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        LogErrorsWithStage logPrefix, "��`�t�@�C���I�[�v��", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    On Error GoTo 0

    If LOG_ENABLED Then
        LogDebug logPrefix & "��`�t�@�C�����J���܂���: " & definitionFilePath
    End If

    currentStage = "GetDefinitionWorksheet"
    On Error Resume Next
    Set definitionWs = definitionWb.Worksheets(definitionSheetName)
    On Error GoTo 0
    If definitionWs Is Nothing Then
        AddError errors, "��`�V�[�g��������܂���: " & definitionSheetName
        SafeCloseWorkbook definitionWb
        LogErrorsWithStage logPrefix, "��`�V�[�g�擾", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "��`�V�[�g���擾���܂���: " & definitionSheetName
    End If

    currentStage = "ReadColumnDefinitions"
    Set columns = ReadColumnDefinitions(definitionWs, colNameAddr, typeAddr, precisionAddr, scaleAddr, pkAddr, notNullAddr, typeCatalog, errors)

    SafeCloseWorkbook definitionWb

    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "��`��ǂݍ���", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "�L���ȃJ������`��1�����擾�ł��܂���B"
        LogErrorsWithStage logPrefix, "��`�񌟏�", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "��`��=" & CStr(columns.Count)
    End If

    If hasDataFile Then
        currentStage = "CheckDataFile"
        If Dir$(dataFilePath, vbNormal) = vbNullString Then
            AddError errors, "�f�[�^�t�@�C���p�X�����݂��܂���: " & dataFilePath
            LogErrorsWithStage logPrefix, "�f�[�^�t�@�C�����݊m�F", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            LogDebug logPrefix & "�f�[�^�t�@�C�����݊m�FOK: " & dataFilePath
        End If

        currentStage = "OpenDataWorkbook"
        On Error Resume Next
        Set dataWb = Application.Workbooks.Open(fileName:=dataFilePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
        If Err.Number <> 0 Then
            AddError errors, "�f�[�^�t�@�C�����J���܂���: " & dataFilePath & " (" & Err.Description & ")"
            Err.Clear
            On Error GoTo 0
            LogErrorsWithStage logPrefix, "�f�[�^�t�@�C���I�[�v��", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        On Error GoTo 0

        If LOG_ENABLED Then
            LogDebug logPrefix & "�f�[�^�t�@�C�����J���܂���: " & dataFilePath
        End If

        currentStage = "GetDataWorksheet"
        On Error Resume Next
        Set dataWs = dataWb.Worksheets(dataSheetName)
        On Error GoTo 0
        If dataWs Is Nothing Then
            AddError errors, "�f�[�^�V�[�g��������܂���: " & dataSheetName
            SafeCloseWorkbook dataWb
            LogErrorsWithStage logPrefix, "�f�[�^�V�[�g�擾", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            LogDebug logPrefix & "�f�[�^�V�[�g���擾���܂���: " & dataSheetName
        End If

        currentStage = "ReadDataRecords"
        Set dataRecords = ReadDataRecords(dataWs, dataStartAddr, columns, errors)

        SafeCloseWorkbook dataWb

        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "�����f�[�^�ǂݍ���", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If

        If LOG_ENABLED Then
            If dataRecords Is Nothing Then
                LogDebug logPrefix & "�����f�[�^����=0 (Nothing)"
            Else
                LogDebug logPrefix & "�����f�[�^����=" & CStr(dataRecords.Count)
            End If
        End If
    End If

    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQL�����J�n"
    End If
    currentStage = "GenerateCreateSql"
    createSql = GenerateCreateSqlText(definitionTableName, columns, dbms, errors)
    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "CREATE SQL����", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQL���������B������=" & CStr(Len(createSql))
    End If

    If hasDataFile Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQL�����J�n"
        End If
        currentStage = "GenerateInsertSql"
        insertSql = GenerateInsertSqlText(dataTableName, columns, dataRecords, dbms, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "INSERT SQL����", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQL���������B������=" & CStr(Len(insertSql))
        End If
    End If

    timestampText = Format$(Now, "yyyymmdd_hhnnss")
    If LOG_ENABLED Then
        LogDebug logPrefix & "CREATE SQL�t�@�C���o�͊J�n"
    End If
    currentStage = "WriteCreateSqlFile"
    createOutputPath = WriteSqlFile(definitionFilePath, SanitizeForFile(tableName:=definitionTableName) & "_CREATE_" & timestampText & ".sql", createSql, errors)
    If errors.Count > 0 Then
        LogErrorsWithStage logPrefix, "CREATE SQL�t�@�C���o��", errors
        WriteErrors mainWs, rowIndex, errors
        Exit Sub
    End If
    If LOG_ENABLED Then
        LogInfo logPrefix & "CREATE SQL�o�͊���: " & createOutputPath
    End If

    If hasDataFile Then
        If LOG_ENABLED Then
            LogDebug logPrefix & "INSERT SQL�t�@�C���o�͊J�n"
        End If
        currentStage = "WriteInsertSqlFile"
        insertOutputPath = WriteSqlFile(definitionFilePath, SanitizeForFile(tableName:=dataTableName) & "_INSERT_" & timestampText & ".sql", insertSql, errors)
        If errors.Count > 0 Then
            LogErrorsWithStage logPrefix, "INSERT SQL�t�@�C���o��", errors
            WriteErrors mainWs, rowIndex, errors
            Exit Sub
        End If
        If LOG_ENABLED Then
            LogInfo logPrefix & "INSERT SQL�o�͊���: " & insertOutputPath
        End If
    End If

    If hasDataFile Then
        mainWs.Cells(rowIndex, MAIN_COL_MESSAGE).value = "SQL�o��: " & createOutputPath & vbLf & insertOutputPath
    Else
        mainWs.Cells(rowIndex, MAIN_COL_MESSAGE).value = "SQL�o��: " & createOutputPath
    End If

    If LOG_ENABLED Then
        If hasDataFile Then
            LogInfo logPrefix & "�������� (CREATE/INSERT)"
        Else
            LogInfo logPrefix & "�������� (CREATE�̂�)"
        End If
    End If

    currentStage = "Finalize"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    On Error GoTo 0
    If LOG_ENABLED Then
        LogErrorMessage logPrefix & "�\�����ʃG���[����(Stage=" & currentStage & "): ErrNumber=" & CStr(Err.Number) & ", Description=" & Err.Description
    End If
    SafeCloseWorkbook definitionWb
    SafeCloseWorkbook dataWb
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
        LogDebug "GenerateInsertSqlText�J�n: table=" & tableName & ", columns=" & columnCountText & ", records=" & recordCountText & ", dbms=" & dbms
    End If

    If columns Is Nothing Or columns.Count = 0 Then
        AddError errors, "�����f�[�^�p�̃J������`�����݂��܂���B"
        Exit Function
    End If

    schemaName = ExtractSchemaName(tableName)
    pureTableName = ExtractTableName(tableName)

    Select Case dbms
        Case "SQLServer"
            If LenB(schemaName) = 0 Then schemaName = "dbo"
        Case "SQLite"
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
            LogInfo "GenerateInsertSqlText: �f�[�^�s�����݂��Ȃ�����INSERT�����o�͂��܂���"
        End If
        GenerateInsertSqlText = "-- �f�[�^�s�����݂��܂���B"
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
                    LogErrorMessage "GenerateInsertSqlText: �l�ϊ��G���[ record=" & CStr(idx) & ", column=" & def.Name & ", address=" & addresses(colIdx)
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
    textValue = CStr(value)
    textValue = Replace(textValue, "'", "''")
    FormatStringLiteral = "'" & textValue & "'"
End Function

Private Function FormatDateLiteral(ByVal value As Variant, ByVal address As String, _
                                   ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "���t�̒l����͂ł��܂���: " & address
        Exit Function
    End If

    Dim dateText As String
    dateText = Format$(dt, "yyyy-mm-dd")

    Select Case dbms
        Case "Oracle"
            FormatDateLiteral = "TO_DATE('" & dateText & "','YYYY-MM-DD')"
        Case "SQLServer"
            FormatDateLiteral = "CAST('" & dateText & "' AS DATETIME)"
        Case Else
            FormatDateLiteral = "'" & dateText & "'"
    End Select
End Function

Private Function FormatTimeLiteral(ByVal value As Variant, ByVal address As String, _
                                   ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "�����̒l����͂ł��܂���: " & address
        Exit Function
    End If

    Dim timeText As String
    timeText = Format$(dt, "HH:nn:ss")

    Select Case dbms
        Case "Oracle"
            FormatTimeLiteral = "TO_TIMESTAMP('" & ORACLE_TIME_BASE_DATE & " " & timeText & "','YYYY-MM-DD HH24:MI:SS')"
        Case "SQLServer"
            FormatTimeLiteral = "CAST('1900-01-01 " & timeText & "' AS DATETIME)"
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
        Case Else
            resolved = normalized
    End Select

    If LenB(resolved) = 0 Then
        AddError errors, "DBMS(" & dbms & ")�ŗ��p�ł��Ȃ��^�ł�: " & rawTypeName
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

Private Function FormatTimestampLiteral(ByVal value As Variant, ByVal address As String, _
                                        ByVal dbms As String, ByVal errors As Collection) As String
    Dim dt As Date
    If Not TryParseDateValue(value, dt) Then
        AddError errors, "�����̒l����͂ł��܂���: " & address
        Exit Function
    End If

    Dim timestampText As String
    timestampText = Format$(dt, "yyyy-mm-dd HH:nn:ss")

    Select Case dbms
        Case "Oracle"
            FormatTimestampLiteral = "TO_TIMESTAMP('" & timestampText & "','YYYY-MM-DD HH24:MI:SS')"
        Case "SQLServer"
            FormatTimestampLiteral = "CAST('" & timestampText & "' AS DATETIME)"
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

    Set nameCell = ResolveSingleCell(targetWs, nameAddr, "�J�������J�n�Z��", errors)
    Set typeCell = ResolveSingleCell(targetWs, typeAddr, "�^�J�����J�n�Z��", errors)
    Set precisionCell = ResolveSingleCell(targetWs, precisionAddr, "�������J�����J�n�Z��", errors)
    Set scaleCell = ResolveSingleCell(targetWs, scaleAddr, "�������J�����J�n�Z��", errors)
    Set pkCell = ResolveSingleCell(targetWs, pkAddr, "PK�J�����J�n�Z��", errors)
    Set notNullCell = ResolveSingleCell(targetWs, notNullAddr, "NotNull�J�����J�n�Z��", errors)

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
            AddError errors, "�^�������͂ł�: " & targetWs.Cells(currentRow, typeCell.Column).address(False, False)
            Exit Do
        End If

        Dim typeKey As String
        typeKey = UCase$(columnType)
        If Not typeCatalog.Exists(typeKey) Then
            AddError errors, "�^��type�V�[�g�ɑ��݂��܂���: " & columnType & " (" & targetWs.Cells(currentRow, typeCell.Column).address(False, False) & ")"
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

        If category = CATEGORY_NUMERIC Or category = CATEGORY_STRING Then
            Dim parsedPrecision As Long
            If Not TryParsePositiveInteger(precisionText, parsedPrecision) Then
                AddError errors, "���������s���ł�: " & targetWs.Cells(currentRow, precisionCell.Column).address(False, False)
            Else
                definition.Precision = parsedPrecision
                definition.HasPrecision = True
            End If
        End If

        If category = CATEGORY_NUMERIC Then
            Dim parsedScale As Long
            If Not TryParseNonNegativeInteger(scaleText, parsedScale) Then
                AddError errors, "���������s���ł�: " & targetWs.Cells(currentRow, scaleCell.Column).address(False, False)
            Else
                definition.ScaleDigits = parsedScale
                definition.HasScale = True
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

    Set startCell = ResolveSingleCell(targetWs, startAddr, "�f�[�^�J�n�Z��", errors)
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
            AddError errors, "No��w�b�_�ʒu�����ł��܂���: " & startCell.address(False, False)
            Exit Function
        End If

        noColumn = LocateNoHeaderColumn(targetWs, headerCandidateRow, startCell.Column)
        If noColumn = 0 Then
            AddError errors, "No��w�b�_��������܂���: " & targetWs.Cells(headerCandidateRow, startCell.Column).address(False, False)
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
            AddError errors, "�����f�[�^��`�w�b�_�ɃJ���������݂��܂���: " & def.Name
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
                    AddError errors, "�K�{���ڂ��󗓂ł�: " & valueCell.address(False, False)
                End If
            ElseIf columns(idx).IsNotNull Then
                If Not HasValueWithoutTrim(cellValue) Then
                    AddError errors, "�K�{���ڂ��󗓂ł�: " & valueCell.address(False, False)
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

Private Function FormatDataType(ByVal definition As ColumnDefinition, ByVal dbms As String, ByVal errors As Collection) As String
    Dim baseType As String
    baseType = ResolveDbmsTypeName(definition.DataType, definition.category, dbms, errors)
    If LenB(baseType) = 0 Or errors.Count > 0 Then Exit Function

    Select Case definition.category
        Case CATEGORY_NUMERIC
            If Not definition.HasPrecision Or Not definition.HasScale Then
                AddError errors, "���l�^�̌���񂪕s�����Ă��܂�: " & definition.Name
                Exit Function
            End If
            FormatDataType = baseType & "(" & CStr(definition.Precision) & "," & CStr(definition.ScaleDigits) & ")"
        Case CATEGORY_STRING
            If Not definition.HasPrecision Then
                AddError errors, "������^�̌���񂪕s�����Ă��܂�: " & definition.Name
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
        Case Else
            BuildDropStatement = "DROP TABLE " & qualifiedName & ";"
    End Select
End Function

Private Function WriteSqlFile(ByVal sourceFilePath As String, ByVal fileName As String, _
                              ByVal sqlText As String, ByVal errors As Collection) As String
    Dim fso As Object
    Dim parentFolder As String
    Dim outputPath As String
    Dim ts As Object

    If LenB(fileName) = 0 Then
        AddError errors, "�o�̓t�@�C�����������������ł��܂���B"
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(sourceFilePath)
    If LenB(parentFolder) = 0 Then
        AddError errors, "�Q�ƃt�@�C���̐e�t�H���_���擾�ł��܂���: " & sourceFilePath
        Exit Function
    End If

    outputPath = fso.BuildPath(parentFolder, fileName)

    On Error Resume Next
    Set ts = fso.CreateTextFile(outputPath, True, False)
    If Err.Number <> 0 Then
        AddError errors, "SQL�t�@�C�����쐬�ł��܂���: " & outputPath & " (" & Err.Description & ")"
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
        AddError errors, label & " �̃Z���Q�Ƃ��s���ł�: " & addressText
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub ValidateRequiredValue(ByVal valueText As String, ByVal label As String, ByVal errors As Collection)
    If LenB(valueText) = 0 Then
        AddError errors, label & " �������͂ł��B"
    End If
End Sub

Private Sub ValidateDbms(ByRef dbms As String, ByVal errors As Collection)
    Dim normalized As String
    normalized = NormalizeDbms(dbms)

    If LenB(normalized) = 0 Then
        AddError errors, "DBMS���s���ł�: " & dbms
    Else
        dbms = normalized
    End If
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
    Dim messageText As String
    messageText = JoinCollection(errors, vbLf)
    mainWs.Cells(rowIndex, MAIN_COL_MESSAGE).value = messageText
    If LOG_ENABLED Then
        Dim flattened As String
        flattened = Replace(messageText, vbLf, " | ")
        LogErrorMessage "Row " & CStr(rowIndex) & " �G���[: " & flattened
    End If
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










