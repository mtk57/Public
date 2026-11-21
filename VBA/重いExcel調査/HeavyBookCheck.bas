Attribute VB_Name = "HeavyBookCheck"
Option Explicit

Private Const VESION As String = "1.1.0"

' 重いExcelブックの調査。対象は保存せず、結果も自動保存しない。
' Windows版 Excel 2019以降。追加の参照設定・外部ライブラリは不要。
' ファイル末尾の文字コード・改行はShift-JIS(CP932)・CRLFとする。

' 判定は「しきい値を超える」とき。0の項目は1件以上で要確認。
Private Const TH_FILE_MB As Double = 50
Private Const TH_OPEN_SEC As Double = 10
Private Const TH_STYLES As Long = 2000
Private Const TH_NAMES As Long = 1000
Private Const TH_EXTRA_ROWS As Long = 1000
Private Const TH_EXTRA_COLS As Long = 100
Private Const TH_FORMULAS As Long = 100000
Private Const TH_VOLATILE As Long = 0
Private Const TH_CF As Long = 100
Private Const TH_VALIDATION As Long = 100000
Private Const TH_SHAPES As Long = 500
Private Const TH_HIDDEN_SHAPES As Long = 0
Private Const TH_TINY_SHAPES As Long = 0
Private Const TH_COMMENTS As Long = 1000
Private Const TH_PIVOTS As Long = 10
Private Const TH_CHARTS As Long = 100
Private Const TH_TABLES As Long = 100
Private Const TH_HYPERLINKS As Long = 1000
Private Const TH_CALC_SEC As Double = 2
Private Const TH_BAD_NAMES As Long = 0
Private Const TH_LINKS As Long = 0
Private Const TH_CACHES As Long = 6
Private Const TH_CONNECTIONS As Long = 0
Private Const TINY_POINTS As Double = 1
Private Const MAX_NAME_ROWS As Long = 10000
Private Const FORMULA_BATCH As Long = 8192
Private Const NG As String = "取得失敗"
Private Const NOT_MEASURED As String = "未計測"
Private Const PROTECTED_NOT_RETRIEVED As String = "保護のため未取得"

' Windows標準API。属性だけを読み、2GBを超えるファイルサイズにも対応する。
Private Type FileAttributeData
    Attributes As Long
    CreationLow As Long
    CreationHigh As Long
    AccessLow As Long
    AccessHigh As Long
    WriteLow As Long
    WriteHigh As Long
    SizeHigh As Long
    SizeLow As Long
End Type
Private Declare PtrSafe Function GetFileAttributesExW Lib "kernel32" ( _
    ByVal fileName As LongPtr, ByVal infoLevel As Long, _
    ByRef attributes As FileAttributeData) As Long

Private Type ExcelSettings
    Calculation As XlCalculation
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    ScreenUpdating As Boolean
    StatusBar As Variant
    AutomationSecurity As Long
End Type

Private mExcel As Excel.Application
Private mTarget As Excel.Workbook
Private mReport As Excel.Workbook
Private mSummary As Excel.Worksheet
Private mSheetReport As Excel.Worksheet
Private mNameReport As Excel.Worksheet
Private mErrorReport As Excel.Worksheet
Private mSummaryRow As Long
Private mErrorCount As Long
Private mErrorRows As Long
Private mLogFailures As Long
Private mErrorCountRow As Long
Private mLogFailuresRow As Long
Private mLastError As String
Private mRunning As Boolean

' ファイルを選択し、別Excelで調査して、結果を表示する。
Public Sub AnalyzeHeavyWorkbook()
    Dim chosen As Variant, doCalc As Boolean, old As ExcelSettings
    Dim settingsSaved As Boolean, aborted As Boolean, targetClosed As Boolean
    Dim started As Double, openSeconds As Double, ws As Excel.Worksheet
    Dim data As Variant, outputRow As Long, i As Long, calcTotal As Variant
    Dim styles As Variant, customStyles As Variant
    Dim number As Long, description As String, message As String
    Dim formatValue As Variant

    If mRunning Then
        MsgBox "調査は既に実行中です。", vbExclamation
        Exit Sub
    End If
    chosen = Application.GetOpenFilename( _
        "Excelブック (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        , "調査するExcelブックを選択")
    If VarType(chosen) = vbBoolean Then Exit Sub
    doCalc = (MsgBox("再計算時間を計測しますか？" & vbCrLf & vbCrLf & _
        "再計算では、外部データ参照が実行され、通信が発生する可能性があります。" & vbCrLf & _
        "調査対象の各シートとブック全体を別インスタンスで計算します。" & vbCrLf & _
        "別インスタンスではアドインを読み込まないため、" & _
        "アドイン関数の計算時間は含まれません。" & vbCrLf & _
        "数式が多い場合は時間がかかります。", _
        vbYesNo Or vbExclamation Or vbDefaultButton2, "再計算の確認") = vbYes)

    mRunning = True
    ResetState
    On Error GoTo Failed
    Set mExcel = New Excel.Application
    ' 引数で空のワークシートを指定し、ユーザーテンプレートを使わない。
    ' CalculationはブックのないExcelでは設定できないため、先に空ブックを作る。
    Set mReport = mExcel.Workbooks.Add(xlWBATWorksheet)
    SaveSettings old
    settingsSaved = True
    mExcel.EnableEvents = False
    mExcel.DisplayAlerts = False
    mExcel.ScreenUpdating = False
    mExcel.AutomationSecurity = msoAutomationSecurityForceDisable
    mExcel.Calculation = xlCalculationManual
    InitializeReport
    AddSummary "調査対象", CStr(chosen), "", "対象・結果とも自動保存しません。"
    AddSummary "調査状態", "調査中", "", "中断時も取得済みの結果を残します。"
    AddSummary "通信について", "接続の更新処理は呼び出しません", "", _
        "開く際の自動接続はExcelの既定セキュリティとブックの設定に従います。"
    AddSummary "ファイルサイズ(MB)", GetFileSizeMB(CStr(chosen)), "自動", _
        "Windowsのファイル属性から取得。1MB=1024×1024バイト。", TH_FILE_MB

    ' パスワードやExcel 4.0マクロ等の確認画面が隠れないよう先に表示する。
    ' 呼び出し元のExcelの設定は変更しない。
    mExcel.Visible = True
    mExcel.StatusBar = "調査対象を読み取り専用で開いています..."
    started = Timer
    Set mTarget = mExcel.Workbooks.Open(Filename:=CStr(chosen), _
        UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, _
        Notify:=False, AddToMru:=False)
    openSeconds = ElapsedSeconds(started)
    If Not mTarget.ReadOnly Then
        Err.Raise vbObjectError + 2001, , "調査対象を読み取り専用で開けませんでした。"
    End If
    If mExcel.Calculation <> xlCalculationManual Then
        Err.Raise vbObjectError + 2002, , "手動計算が維持されていません。調査を中断します。"
    End If
    AddSummary "開く時間(秒)", Round(openSeconds, 3), "自動", _
        "読み取り専用・マクロ無効・リンク更新なし・手動計算での時間。", TH_OPEN_SEC
    formatValue = BookMetric(mTarget, 1)
    AddSummary "ファイル形式", formatValue, FormatJudgment(CStr(chosen), formatValue), _
        ".xlsは旧形式です。形式変換の要否を確認してください。調査中に変換はしません。"
    AddSummary "ワークシート数", BookMetric(mTarget, 2), "", "グラフシート等はシート別の対象外です。"
    CountStyles mTarget, styles, customStyles
    AddSummary "セルスタイル数", styles, "自動", "", TH_STYLES
    AddSummary "うちユーザー定義スタイル数", customStyles, "", ""
    AnalyzeNames mTarget
    AnalyzeLinks mTarget, xlExcelLinks, "Excel外部リンク"
    AnalyzeLinks mTarget, xlOLELinks, "OLE外部リンク"
    AddSummary "ピボットキャッシュ数", BookMetric(mTarget, 3), "自動", "", TH_CACHES
    AnalyzeConnections mTarget

    outputRow = 2
    For Each ws In mTarget.Worksheets
        mExcel.StatusBar = "シートを調査中: " & ws.Name
        data = AnalyzeSheet(ws, doCalc)
        WriteRow mSheetReport, outputRow, data
        outputRow = outputRow + 1
    Next ws
    If Not doCalc Then
        AddSummary "再計算時間合計(秒)", NOT_MEASURED, "", _
            "利用者が許可していないため未実行。別インスタンスではアドインを読み込まないため、" & _
            "アドイン関数の計算時間は含まれません。"
    Else
        calcTotal = CalculateFullWorkbook()
        AddSummary "再計算時間合計(秒)", calcTotal, "自動", _
            "全シートの計測後に別インスタンスのCalculateFullを計測。シート別の加算値ではありません。" & _
            "別インスタンスではアドインを読み込まないため、アドイン関数の計算時間は含まれません。", TH_CALC_SEC
    End If
    GoTo Finish

Failed:
    number = Err.Number
    description = Err.Description
    aborted = True
    RecordError "マクロ全体", "調査を中断", number, description
    Resume Finish

Finish:
    On Error Resume Next
    ' Failedへの再入を防ぐ。個々の後始末は独立したハンドラーでエラーを記録する。
    targetClosed = CloseTarget()
    If settingsSaved Then
        For i = 1 To 6
            If i = 2 And Not targetClosed Then
                ' 対象を閉じられない異常時は、自動計算への復元で未許可計算を起こさない。
                RecordError "後始末", "Calculationの復元", vbObjectError + 2003, _
                    "対象を閉じられないため手動計算を維持しました。表示されたExcelを確認してください。"
            Else
                RestoreSetting old, i
            End If
        Next i
    End If
    FinalizeReport aborted
    If Not ShowResults() Then QuitOwnedExcel
    UpdateErrorCounts
    message = "調査が完了しました。"
    If aborted Then message = "調査を中断しました。取得済みの結果を確認してください。"
    If mErrorCount > 0 Then
        message = message & vbCrLf & "取得・処理に失敗した記録: " & mErrorCount & " 件" & _
            vbCrLf & "「エラー記録」シートを確認してください。"
    End If
    If mLogFailures > 0 Then
        message = message & vbCrLf & "エラー記録への出力失敗: " & mLogFailures & " 件"
    End If
    If mLastError <> "" And (aborted Or mLogFailures > 0) Then
        message = message & vbCrLf & "最後のエラー: " & mLastError
    End If
    Set mTarget = Nothing
    Set mSummary = Nothing
    Set mSheetReport = Nothing
    Set mNameReport = Nothing
    Set mErrorReport = Nothing
    Set mReport = Nothing
    Set mExcel = Nothing
    mRunning = False
    If aborted Or mErrorCount > 0 Then
        MsgBox message, vbExclamation, "重いExcelの調査"
    Else
        MsgBox message & vbCrLf & "判定列の「要確認」を確認してください。", vbInformation, "重いExcelの調査"
    End If
End Sub

' 前回の実行状態を解放し、今回のカウンターを初期化する。
Private Sub ResetState()
    Set mTarget = Nothing
    Set mReport = Nothing
    Set mSummary = Nothing
    Set mSheetReport = Nothing
    Set mNameReport = Nothing
    Set mErrorReport = Nothing
    Set mExcel = Nothing
    mErrorCount = 0
    mErrorRows = 0
    mLogFailures = 0
    mErrorCountRow = 0
    mLogFailuresRow = 0
    mLastError = ""
    mSummaryRow = 2
End Sub

' 別インスタンスの変更前の設定を保存する。
Private Sub SaveSettings(ByRef old As ExcelSettings)
    With mExcel
        old.Calculation = .Calculation
        old.EnableEvents = .EnableEvents
        old.DisplayAlerts = .DisplayAlerts
        old.ScreenUpdating = .ScreenUpdating
        old.StatusBar = .StatusBar
        old.AutomationSecurity = .AutomationSecurity
    End With
End Sub

' エラー記録シートを最初に作り、ほかの出力シートを準備する。
Private Sub InitializeReport()
    Set mErrorReport = mReport.Worksheets(1)
    mErrorReport.Name = "エラー記録"
    WriteRow mErrorReport, 1, Array("場所", "項目", "エラー番号", "内容")
    Set mSummary = mReport.Worksheets.Add(Before:=mErrorReport)
    mSummary.Name = "概要"
    Set mSheetReport = mReport.Worksheets.Add(Before:=mErrorReport)
    mSheetReport.Name = "シート別"
    Set mNameReport = mReport.Worksheets.Add(Before:=mErrorReport)
    mNameReport.Name = "問題のある名前定義"
    WriteRow mSummary, 1, Array("項目", "値", "判定", "補足")
    WriteRow mNameReport, 1, Array("名前", "参照先", "表示状態", "問題")
    WriteRow mSheetReport, 1, Array("シート名", "表示状態", "使用範囲の最終セル(Ctrl+End)", _
        "実データの最終セル", "余分な行数", "余分な列数", "数式セル数", "揮発性関数を含むセル数", _
        "INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN", "CELL", "INFO", _
        "条件付き書式数", "入力規則セル数", "図形数", "うち非表示", "うち1pt未満", _
        "コメント合計", "従来のコメント(メモ)", "スレッドコメント", "ピボットテーブル数", _
        "グラフ数", "テーブル数", "ハイパーリンク数", "再計算秒", "判定", "判定の理由")
End Sub

' 値を行単位で書く。文字列は数式・リンクとして評価させない。
Private Sub WriteRow(ByVal ws As Excel.Worksheet, ByVal rowNumber As Long, ByVal values As Variant)
    Dim buffer() As Variant, i As Long, count As Long, value As Variant
    count = UBound(values) - LBound(values) + 1
    ReDim buffer(1 To 1, 1 To count)
    For i = 1 To count
        value = values(LBound(values) + i - 1)
        If VarType(value) = vbString Then
            value = Left$(CStr(value), 32767)
            ws.Cells(rowNumber, i).NumberFormat = "@"
        Else
            ws.Cells(rowNumber, i).NumberFormat = "General"
        End If
        buffer(1, i) = value
    Next i
    ws.Cells(rowNumber, 1).Resize(1, count).Value2 = buffer
End Sub

' 概要に1行追加する。「自動」の指定時は数値と失敗値から判定する。
Private Sub AddSummary(ByVal item As String, ByVal value As Variant, ByVal judgment As String, _
    ByVal note As String, Optional ByVal threshold As Double = 0)
    If judgment = "自動" Then judgment = JudgeNumber(value, threshold)
    If IsFailure(value) Then judgment = "要確認"
    WriteRow mSummary, mSummaryRow, Array(item, value, judgment, note)
    mSummaryRow = mSummaryRow + 1
End Sub

' エラーを即時出力する。記録自体が失敗しても後始末を止めない。
Private Sub RecordError(ByVal place As String, ByVal item As String, _
    ByVal number As Long, ByVal description As String)
    On Error GoTo Failed
    mErrorCount = mErrorCount + 1
    mLastError = place & " / " & item & " / " & CStr(number) & " / " & description
    If mErrorReport Is Nothing Then
        mLogFailures = mLogFailures + 1
        Exit Sub
    End If
    WriteRow mErrorReport, mErrorRows + 2, Array(place, item, number, description)
    mErrorRows = mErrorRows + 1
    Exit Sub
Failed:
    mLogFailures = mLogFailures + 1
End Sub

' ファイル属性から符号なし64ビットのサイズを組み立てる。書き込みは行わない。
Private Function GetFileSizeMB(ByVal path As String) As Variant
    Dim info As FileAttributeData, result As Long, number As Long
    On Error GoTo Failed
    result = GetFileAttributesExW(StrPtr(path), 0, info)
    If result = 0 Then
        number = Err.LastDllError
        RecordError "ブック", "ファイルサイズ", number, "ファイル属性を取得できませんでした。"
        GetFileSizeMB = NG
    Else
        GetFileSizeMB = (UnsignedLong(info.SizeHigh) * 4294967296# + _
            UnsignedLong(info.SizeLow)) / 1048576#
    End If
    Exit Function
Failed:
    RecordError "ブック", "ファイルサイズ", Err.Number, Err.Description
    GetFileSizeMB = NG
End Function

' APIの32ビット値を符号なしのDoubleとして解釈する。
Private Function UnsignedLong(ByVal value As Long) As Double
    UnsignedLong = CDbl(value)
    If value < 0 Then UnsignedLong = UnsignedLong + 4294967296#
End Function

' ブックの基本プロパティを個別に取得する。
Private Function BookMetric(ByVal wb As Excel.Workbook, ByVal metric As Long) As Variant
    On Error GoTo Failed
    Select Case metric
        Case 1: BookMetric = wb.FileFormat
        Case 2: BookMetric = wb.Worksheets.Count
        Case 3: BookMetric = wb.PivotCaches.Count
        Case Else: Err.Raise 5, , "未知のブック調査項目です。"
    End Select
    Exit Function
Failed:
    RecordError "ブック", "基本項目 " & CStr(metric), Err.Number, Err.Description
    BookMetric = NG
End Function

' スタイル総数とユーザー定義数を分けて取得する。
Private Sub CountStyles(ByVal wb As Excel.Workbook, ByRef total As Variant, ByRef custom As Variant)
    Dim st As Excel.Style, count As Long
    total = GetStyleTotal(wb)
    On Error GoTo Failed
    For Each st In wb.Styles
        If Not st.BuiltIn Then count = count + 1
    Next st
    custom = count
    Exit Sub
Failed:
    RecordError "ブック", "ユーザー定義スタイル数", Err.Number, Err.Description
    custom = NG
End Sub

' 総スタイル数は個々のBuiltInの取得失敗から独立させる。
Private Function GetStyleTotal(ByVal wb As Excel.Workbook) As Variant
    On Error GoTo Failed
    GetStyleTotal = wb.Styles.Count
    Exit Function
Failed:
    RecordError "ブック", "セルスタイル数", Err.Number, Err.Description
    GetStyleTotal = NG
End Function

' 拡張子またはExcelの形式値が旧形式なら注意喚起する。
Private Function FormatJudgment(ByVal path As String, ByVal fileFormat As Variant) As String
    If LCase$(Right$(path, 4)) = ".xls" Then FormatJudgment = "要確認"
    If IsNumeric(fileFormat) Then
        If CLng(fileFormat) = xlExcel8 Or CLng(fileFormat) = xlExcel9795 Then FormatJudgment = "要確認"
    ElseIf IsFailure(fileFormat) Then
        FormatJudgment = "要確認"
    End If
End Function

' 数値がしきい値を超えるか、取得に失敗したときに要確認を返す。
Private Function JudgeNumber(ByVal value As Variant, ByVal threshold As Double) As String
    If IsFailure(value) Then
        JudgeNumber = "要確認"
    ElseIf IsNumeric(value) Then
        If CDbl(value) > threshold Then JudgeNumber = "要確認"
    End If
End Function

' Variantの数値と文字列を混同せず、失敗の表示値を判定する。
Private Function IsFailure(ByVal value As Variant) As Boolean
    If VarType(value) = vbString Then IsFailure = (CStr(value) = NG)
End Function

' Timerの日付またぎを補正する。24時間未満の計測を対象とする。
Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = CDbl(Timer) - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

' ワークシート1枚分の項目を独立して取得する。配列は0始まり。
Private Function AnalyzeSheet(ByVal ws As Excel.Worksheet, ByVal doCalc As Boolean) As Variant
    Dim values(0 To 30) As Variant, formulas As Excel.Range, counts As Variant
    Dim i As Long, total As Variant, hidden As Variant, tiny As Variant
    Dim reason As String
    values(0) = ws.Name
    values(1) = SheetVisibility(ws)
    GetExtent ws, values(2), values(3), values(4), values(5)
    GetFormulas ws, formulas, values(6)
    counts = CountVolatile(formulas, values(6), ws.Name)
    For i = 0 To 8
        values(7 + i) = counts(i)
    Next i
    values(16) = SheetMetric(ws, 1)
    values(17) = CountValidation(ws)
    CountShapes ws, total, hidden, tiny
    values(18) = total
    values(19) = hidden
    values(20) = tiny
    values(21) = SheetMetric(ws, 2)
    values(23) = SheetMetric(ws, 3)
    If IsNumeric(values(21)) And IsNumeric(values(23)) Then
        If CDbl(values(21)) >= CDbl(values(23)) Then
            values(22) = CDbl(values(21)) - CDbl(values(23))
        Else
            RecordError ws.Name, "コメントの内訳", vbObjectError + 2004, _
                "コメント合計よりスレッドコメント数が大きく、メモ数を確定できません。"
            values(22) = NG
        End If
    ElseIf CStr(values(23)) = "非対応" Then
        values(22) = "非対応"
    Else
        values(22) = NG
    End If
    For i = 4 To 7
        values(20 + i) = SheetMetric(ws, i)
    Next i
    values(28) = NOT_MEASURED
    If doCalc Then values(28) = CalculateSheet(ws)
    AddReason reason, values(4), TH_EXTRA_ROWS, "余分な行"
    AddReason reason, values(5), TH_EXTRA_COLS, "余分な列"
    AddReason reason, values(6), TH_FORMULAS, "数式"
    AddReason reason, values(7), TH_VOLATILE, "揮発性関数"
    AddReason reason, values(16), TH_CF, "条件付き書式"
    AddReason reason, values(17), TH_VALIDATION, "入力規則"
    If CStr(values(17)) = PROTECTED_NOT_RETRIEVED Then
        AppendReason reason, "入力規則：" & PROTECTED_NOT_RETRIEVED
    End If
    AddReason reason, values(18), TH_SHAPES, "図形"
    AddReason reason, values(19), TH_HIDDEN_SHAPES, "非表示の図形"
    AddReason reason, values(20), TH_TINY_SHAPES, "微小な図形"
    AddReason reason, values(21), TH_COMMENTS, "コメント"
    AddReason reason, values(24), TH_PIVOTS, "ピボットテーブル"
    AddReason reason, values(25), TH_CHARTS, "グラフ"
    AddReason reason, values(26), TH_TABLES, "テーブル"
    AddReason reason, values(27), TH_HYPERLINKS, "ハイパーリンク"
    AddReason reason, values(28), TH_CALC_SEC, "再計算"
    For i = 1 To 28
        If IsFailure(values(i)) Then
            AppendReason reason, "取得失敗あり"
            Exit For
        End If
    Next i
    If reason <> "" Then values(29) = "要確認"
    values(30) = reason
    AnalyzeSheet = values
End Function

' シートの表示状態を取得する。
Private Function SheetVisibility(ByVal ws As Excel.Worksheet) As Variant
    On Error GoTo Failed
    Select Case ws.Visible
        Case xlSheetVisible: SheetVisibility = "表示"
        Case xlSheetHidden: SheetVisibility = "非表示"
        Case xlSheetVeryHidden: SheetVisibility = "完全非表示"
    End Select
    Exit Function
Failed:
    RecordError ws.Name, "シートの表示状態", Err.Number, Err.Description
    SheetVisibility = NG
End Function

' Ctrl+End相当をUsedRange取得より先に読み、実データの矩形の右下と比較する。
Private Sub GetExtent(ByVal ws As Excel.Worksheet, ByRef usedAddress As Variant, _
    ByRef dataAddress As Variant, ByRef extraRows As Variant, ByRef extraCols As Variant)
    Dim last As Excel.Range, lastRow As Variant, lastCol As Variant, used As Excel.Range
    Dim usedRow As Long, usedCol As Long, usedOK As Boolean
    On Error GoTo UsedFailed
    If IsProtectedWorksheet(ws) Then
        Set used = ws.UsedRange
        Set last = used.Cells(used.Rows.Count, used.Columns.Count)
    Else
        Set last = ws.Cells.SpecialCells(xlCellTypeLastCell)
    End If
    usedRow = last.Row
    usedCol = last.Column
    usedAddress = last.Address(False, False, xlA1)
    usedOK = True
    GoTo DataPart
UsedFailed:
    RecordError ws.Name, "使用範囲の最終セル", Err.Number, Err.Description
    usedAddress = NG
    Resume DataPart
DataPart:
    On Error GoTo DataFailed
    lastRow = LastDataIndex(ws, xlByRows)
    lastCol = LastDataIndex(ws, xlByColumns)
    If IsFailure(lastRow) Or IsFailure(lastCol) Then
        dataAddress = NG
        extraRows = NG
        extraCols = NG
    Else
        If CDbl(lastRow) = 0 And CDbl(lastCol) = 0 Then
            dataAddress = "(実データなし)"
            ' 空シートはA1を基準にし、書式もない空シートは余分0とする。
            lastRow = 1
            lastCol = 1
        Else
            dataAddress = ws.Cells(CLng(lastRow), CLng(lastCol)).Address(False, False, xlA1)
        End If
        If usedOK Then
            extraRows = usedRow - CLng(lastRow)
            extraCols = usedCol - CLng(lastCol)
            If extraRows < 0 Then extraRows = 0
            If extraCols < 0 Then extraCols = 0
        Else
            extraRows = NG
            extraCols = NG
        End If
    End If
    Exit Sub
DataFailed:
    RecordError ws.Name, "実データ範囲・余分な行列", Err.Number, Err.Description
    dataAddress = NG
    extraRows = NG
    extraCols = NG
End Sub

' 書式だけのセルを除外し、数式（結果が空文字でも含む）と値の最終位置を探す。
Private Function LastDataIndex(ByVal ws As Excel.Worksheet, ByVal order As XlSearchOrder) As Variant
    Dim found As Excel.Range
    On Error GoTo Failed
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=order, _
        SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If found Is Nothing Then
        LastDataIndex = 0
    ElseIf order = xlByRows Then
        LastDataIndex = found.Row
    Else
        LastDataIndex = found.Column
    End If
    Exit Function
Failed:
    RecordError ws.Name, "実データの最終位置 " & CStr(order), Err.Number, Err.Description
    LastDataIndex = NG
End Function

' 数式の存在を先に調べ、SpecialCellsの1004を一律0とすることを避ける。
Private Sub GetFormulas(ByVal ws As Excel.Worksheet, ByRef formulas As Excel.Range, ByRef count As Variant)
    Dim used As Excel.Range, has As Variant
    On Error GoTo Failed
    Set used = ws.UsedRange
    has = used.HasFormula
    If Not IsNull(has) Then
        If has = False Then
            count = 0
            Exit Sub
        End If
    End If
    If IsProtectedWorksheet(ws) Then
        CollectProtectedFormulas used, formulas
    Else
        Set formulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
    End If
    If formulas Is Nothing Then
        count = 0
    Else
        count = formulas.CountLarge
    End If
    Exit Sub
Failed:
    RecordError ws.Name, "数式セル数", Err.Number, Err.Description
    Set formulas = Nothing
    count = NG
End Sub

' 保護中はSpecialCellsを使わない読取り経路にする。
Private Function IsProtectedWorksheet(ByVal ws As Excel.Worksheet) As Boolean
    IsProtectedWorksheet = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
End Function

' HasFormulaが混在する範囲だけを分割し、保護された数式セルを変更せず集約する。
Private Sub CollectProtectedFormulas(ByVal part As Excel.Range, ByRef formulas As Excel.Range)
    Dim has As Variant, half As Long, columns As Long, rows As Long
    has = part.HasFormula
    If Not IsNull(has) Then
        If CBool(has) Then
            If formulas Is Nothing Then
                Set formulas = part
            Else
                Set formulas = part.Application.Union(formulas, part)
            End If
        End If
        Exit Sub
    End If
    columns = part.Columns.Count
    rows = part.Rows.Count
    If columns > 1 Then
        half = columns \ 2
        CollectProtectedFormulas part.Resize(rows, half), formulas
        CollectProtectedFormulas part.Cells(1, half + 1).Resize(rows, columns - half), formulas
    ElseIf rows > 1 Then
        half = rows \ 2
        CollectProtectedFormulas part.Resize(half, columns), formulas
        CollectProtectedFormulas part.Cells(half + 1, 1).Resize(rows - half, columns), formulas
    Else
        Err.Raise vbObjectError + 2008, , "単一セルの数式有無を確定できませんでした。"
    End If
End Sub

' 数式を配列で分割取得し、セル内の関数ごとの重複と関数間の重複を分けて集計する。
Private Function CountVolatile(ByVal formulas As Excel.Range, ByVal formulaCount As Variant, _
    ByVal place As String) As Variant
    Dim counts(0 To 8) As Variant, totals(0 To 8) As Double
    Dim area As Excel.Range, block As Excel.Range, buffer As Variant
    Dim startRow As Long, rows As Long, batchRows As Long, columns As Long
    Dim r As Long, c As Long, i As Long, formulaText As String
    If IsFailure(formulaCount) Then
        For i = 0 To 8: counts(i) = NG: Next i
        CountVolatile = counts
        Exit Function
    End If
    On Error GoTo Failed
    If Not formulas Is Nothing Then
        For Each area In formulas.Areas
            columns = area.Columns.Count
            batchRows = FORMULA_BATCH \ columns
            If batchRows < 1 Then batchRows = 1
            startRow = 1
            Do While startRow <= area.Rows.Count
                rows = area.Rows.Count - startRow + 1
                If rows > batchRows Then rows = batchRows
                Set block = area.Cells(startRow, 1).Resize(rows, columns)
                buffer = block.Formula
                If block.CountLarge = 1 Then
                    formulaText = UCase$(CStr(buffer))
                    If HasVolatileName(formulaText) Then CountFormulaText formulaText, totals
                Else
                    For r = 1 To rows
                        For c = 1 To columns
                            formulaText = UCase$(CStr(buffer(r, c)))
                            If HasVolatileName(formulaText) Then CountFormulaText formulaText, totals
                        Next c
                    Next r
                End If
                startRow = startRow + rows
            Loop
        Next area
    End If
    For i = 0 To 8: counts(i) = totals(i): Next i
    CountVolatile = counts
    Exit Function
Failed:
    RecordError place, "揮発性関数の集計", Err.Number, Err.Description
    For i = 0 To 8: counts(i) = NG: Next i
    CountVolatile = counts
End Function

' 大文字化した数式をInStrで事前に絞り込み、対象関数がない場合の字句解析を省く。
Private Function HasVolatileName(ByVal upperFormula As String) As Boolean
    HasVolatileName = (InStr(1, upperFormula, "INDIRECT", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "OFFSET", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "TODAY", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "NOW", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "RAND", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "RANDBETWEEN", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "CELL", vbBinaryCompare) > 0 Or _
        InStr(1, upperFormula, "INFO", vbBinaryCompare) > 0)
End Function

' 数式の文字列・引用した参照名・構造化参照を飛ばし、関数呼び出しを字句で判定する。
Private Sub CountFormulaText(ByVal formula As String, ByRef totals() As Double)
    Dim found(1 To 8) As Boolean, anyFound As Boolean
    Dim i As Long, j As Long, first As Long, length As Long, index As Long
    Dim ch As String, token As String
    length = Len(formula)
    i = 1
    Do While i <= length
        ch = Mid$(formula, i, 1)
        If ch = """" Or ch = "'" Then
            i = SkipQuoted(formula, i, ch)
        ElseIf ch = "[" Then
            i = SkipBrackets(formula, i)
        ElseIf IsIdentifierChar(ch) Then
            first = i
            Do While i <= length
                If Not IsIdentifierChar(Mid$(formula, i, 1)) Then Exit Do
                i = i + 1
            Loop
            token = UCase$(Mid$(formula, first, i - first))
            j = i
            Do While j <= length
                ch = Mid$(formula, j, 1)
                If ch <> " " And ch <> vbTab And ch <> vbLf And ch <> vbCr Then Exit Do
                j = j + 1
            Loop
            If j <= length Then
                If Mid$(formula, j, 1) = "(" Then
                    index = VolatileIndex(token)
                    If index > 0 Then found(index) = True
                End If
            End If
        Else
            i = i + 1
        End If
    Loop
    For index = 1 To 8
        If found(index) Then
            totals(index) = totals(index) + 1
            anyFound = True
        End If
    Next index
    If anyFound Then totals(0) = totals(0) + 1
End Sub

' エスケープされた同じ引用符（""または''）を含む引用範囲を飛ばす。
Private Function SkipQuoted(ByVal text As String, ByVal first As Long, ByVal quote As String) As Long
    Dim i As Long
    i = first + 1
    Do While i <= Len(text)
        If Mid$(text, i, 1) = quote Then
            If Mid$(text, i + 1, 1) = quote Then
                i = i + 2
            Else
                SkipQuoted = i + 1
                Exit Function
            End If
        Else
            i = i + 1
        End If
    Loop
    SkipQuoted = Len(text) + 1
End Function

' 角括弧内の列名・ブック名を関数名として数えない。
Private Function SkipBrackets(ByVal text As String, ByVal first As Long) As Long
    Dim i As Long, depth As Long, ch As String
    i = first
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If ch = "'" Then
            ' 構造化参照中のアポストロフィは直後の特殊文字をエスケープする。
            i = i + 2
        Else
            If ch = "[" Then depth = depth + 1
            If ch = "]" Then depth = depth - 1
            i = i + 1
            If depth = 0 Then Exit Do
        End If
    Loop
    SkipBrackets = i
End Function

' 識別子の一部として扱う文字を判定し、MYOFFSETや日本語付き名前の誤検出を防ぐ。
Private Function IsIdentifierChar(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) = 0 Then Exit Function
    code = AscW(ch)
    IsIdentifierChar = (ch Like "[A-Za-z0-9_]" Or ch = "." Or ch = "\" Or code < 0 Or code >= 128)
End Function

' Excelの既知の互換プレフィックスだけを取り除いて関数番号に変換する。
Private Function VolatileIndex(ByVal token As String) As Long
    Do
        If Left$(token, 6) = "_XLFN." Or Left$(token, 6) = "_XLWS." Then
            token = Mid$(token, 7)
        Else
            Exit Do
        End If
    Loop
    Select Case token
        Case "INDIRECT": VolatileIndex = 1
        Case "OFFSET": VolatileIndex = 2
        Case "TODAY": VolatileIndex = 3
        Case "NOW": VolatileIndex = 4
        Case "RAND": VolatileIndex = 5
        Case "RANDBETWEEN": VolatileIndex = 6
        Case "CELL": VolatileIndex = 7
        Case "INFO": VolatileIndex = 8
    End Select
End Function

' 保護中は入力規則を取得しない。非保護時は「該当なし」だけを正常な0件とする。
Private Function CountValidation(ByVal ws As Excel.Worksheet) As Variant
    Dim cells As Excel.Range, querying As Boolean, number As Long, description As String
    On Error GoTo Failed
    If IsProtectedWorksheet(ws) Then
        CountValidation = PROTECTED_NOT_RETRIEVED
        Exit Function
    End If
    querying = True
    Set cells = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    querying = False
    CountValidation = cells.CountLarge
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    If querying And number = 1004 And IsNoCellsMessage(description) Then
        CountValidation = 0
    Else
        RecordError ws.Name, "入力規則セル数", number, description
        CountValidation = NG
    End If
End Function

' 日本語・英語の「該当なし」だけを認め、ほかの1004を誤って0にしない。
Private Function IsNoCellsMessage(ByVal description As String) As Boolean
    Dim text As String
    text = LCase$(Trim$(description))
    IsNoCellsMessage = (text = "no cells were found." Or text = "no cells were found" Or _
        text = "該当するセルが見つかりません。" Or text = "該当するセルが見つかりません")
End Function

' コメント枠を除くトップレベルの図形と、その非表示・微小サイズの件数を数える。
Private Sub CountShapes(ByVal ws As Excel.Worksheet, ByRef total As Variant, _
    ByRef hidden As Variant, ByRef tiny As Variant)
    Dim shp As Excel.Shape, n As Long, h As Long, z As Long
    On Error GoTo Failed
    For Each shp In ws.Shapes
        If shp.Type <> msoComment Then
            n = n + 1
            If shp.Visible = msoFalse Then h = h + 1
            If shp.Width < TINY_POINTS Or shp.Height < TINY_POINTS Then z = z + 1
        End If
    Next shp
    total = n
    hidden = h
    tiny = z
    Exit Sub
Failed:
    RecordError ws.Name, "図形数・非表示・微小サイズ", Err.Number, Err.Description
    total = NG
    hidden = NG
    tiny = NG
End Sub

' 各コレクションを直接取得し、非対応や取得失敗を0件として扱わない。
Private Function SheetMetric(ByVal ws As Excel.Worksheet, ByVal metric As Long) As Variant
    On Error GoTo Failed
    Select Case metric
        Case 1: SheetMetric = ws.Cells.FormatConditions.Count
        Case 2: SheetMetric = ws.Comments.Count
        Case 3: SheetMetric = CountThreadedComments(ws)
        Case 4: SheetMetric = ws.PivotTables.Count
        Case 5: SheetMetric = ws.ChartObjects.Count
        Case 6: SheetMetric = ws.ListObjects.Count
        Case 7: SheetMetric = ws.Hyperlinks.Count
        Case Else: Err.Raise 5, , "未知のシート調査項目です。"
    End Select
    Exit Function
Failed:
    RecordError ws.Name, SheetMetricLabel(metric), Err.Number, Err.Description
    SheetMetric = NG
End Function

' スレッドコメントAPIがないExcelでもコンパイルし、非対応は通常の表示として返す。
Private Function CountThreadedComments(ByVal ws As Object) As Variant
    Dim number As Long, description As String
    On Error GoTo Failed
    CountThreadedComments = ws.CommentsThreaded.Count
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    If number = 438 Then
        CountThreadedComments = "非対応"
    Else
        RecordError ws.Name, "スレッドコメント数", number, description
        CountThreadedComments = NG
    End If
End Function

' エラー記録用の調査項目名を返す。
Private Function SheetMetricLabel(ByVal metric As Long) As String
    Select Case metric
        Case 1: SheetMetricLabel = "条件付き書式数"
        Case 2: SheetMetricLabel = "コメント合計"
        Case 3: SheetMetricLabel = "スレッドコメント数"
        Case 4: SheetMetricLabel = "ピボットテーブル数"
        Case 5: SheetMetricLabel = "グラフ数"
        Case 6: SheetMetricLabel = "テーブル数"
        Case 7: SheetMetricLabel = "ハイパーリンク数"
        Case Else: SheetMetricLabel = "不明な項目"
    End Select
End Function

' 許可された場合にだけ、指定シートの使用範囲を計算して時間を測る。
Private Function CalculateSheet(ByVal ws As Excel.Worksheet) As Variant
    Dim started As Double
    On Error GoTo Failed
    If Not ws.EnableCalculation Then
        Err.Raise vbObjectError + 2007, , "シートのEnableCalculationがFalseのため計測できません。設定は変更しません。"
    End If
    started = Timer
    ws.UsedRange.Calculate
    CalculateSheet = Round(ElapsedSeconds(started), 3)
    Exit Function
Failed:
    RecordError ws.Name, "再計算時間", Err.Number, Err.Description
    CalculateSheet = NG
End Function

' 全シートの個別計測後、専用Excel全体の完全再計算時間を別途測る。
Private Function CalculateFullWorkbook() As Variant
    Dim started As Double
    On Error GoTo Failed
    started = Timer
    mExcel.CalculateFull
    CalculateFullWorkbook = Round(ElapsedSeconds(started), 3)
    Exit Function
Failed:
    RecordError "ブック", "再計算時間合計", Err.Number, Err.Description
    CalculateFullWorkbook = NG
End Function

' 判定理由にしきい値を超えた項目を追加する。
Private Sub AddReason(ByRef reason As String, ByVal value As Variant, _
    ByVal threshold As Double, ByVal label As String)
    If IsNumeric(value) Then
        If CDbl(value) > threshold Then AppendReason reason, label
    End If
End Sub

' 理由の区切りを統一する。
Private Sub AppendReason(ByRef reason As String, ByVal label As String)
    If reason <> "" Then reason = reason & "、"
    reason = reason & label
End Sub

' 全名前定義を調べ、問題のある名前だけを最大1万件出力する。総数の集計は打ち切らない。
Private Sub AnalyzeNames(ByVal wb As Excel.Workbook)
    Dim nm As Excel.Name, total As Variant, nameText As Variant, refers As Variant, visible As Variant
    Dim broken As Long, external As Long, hidden As Long, failed As Long
    Dim brokenFailed As Boolean, externalFailed As Boolean, hiddenFailed As Boolean
    Dim problems As Long, shown As Long, reason As String, display As Variant
    Dim rowFailed As Boolean, number As Long, description As String
    Dim index As Long, context As String
    Dim nameOK As Boolean, refersOK As Boolean, visibleOK As Boolean
    On Error GoTo Failed
    total = wb.Names.Count
    For Each nm In wb.Names
        index = index + 1
        context = "名前定義 " & CStr(index)
        nameText = NameProperty(nm, 1, context, nameOK)
        If nameOK Then context = context & " / " & CStr(nameText)
        refers = NameProperty(nm, 2, context, refersOK)
        visible = NameProperty(nm, 3, context, visibleOK)
        reason = ""
        rowFailed = False
        If Not refersOK Then
            brokenFailed = True
            externalFailed = True
            rowFailed = True
        Else
            If HasBrokenReference(CStr(refers)) Then
                broken = broken + 1
                AppendReason reason, "#REF!"
            End If
            If HasExternalReference(CStr(refers)) Then
                external = external + 1
                AppendReason reason, "外部参照"
            End If
        End If
        display = NG
        If Not visibleOK Or Not nameOK Then
            hiddenFailed = True
            rowFailed = True
        ElseIf CBool(visible) Then
            display = "表示"
        Else
            display = "非表示"
            If Not IsInternalName(CStr(nameText)) Then
                hidden = hidden + 1
                AppendReason reason, "非表示"
            End If
        End If
        If rowFailed Then
            failed = failed + 1
            AppendReason reason, NG
        End If
        If reason <> "" Then
            problems = problems + 1
            If shown < MAX_NAME_ROWS Then
                WriteRow mNameReport, shown + 2, Array(nameText, refers, display, reason)
                shown = shown + 1
            End If
        End If
    Next nm
    AddSummary "名前定義数", total, "自動", "シートローカル・内部名も含む総数。", TH_NAMES
    AddSummary "うち #REF! の名前定義", CompleteCount(broken, brokenFailed), "自動", _
        "文字列・引用した識別子内の #REF! は除外。", TH_BAD_NAMES
    AddSummary "うち外部参照の名前定義", CompleteCount(external, externalFailed), "自動", _
        "外部ブックへの参照を集計。構造化参照の角括弧だけでは判定しません。", TH_BAD_NAMES
    AddSummary "うち非表示の名前定義", CompleteCount(hidden, hiddenFailed), "自動", _
        "_xlfn. / _xlpm. / _xleta. / _xlnm. / _FilterDatabase は非表示の問題から除外。", TH_BAD_NAMES
    AddSummary "取得に失敗した名前定義数", failed, "自動", "いずれかのプロパティを取得できなかった名前数。", 0
    AddSummary "問題のある名前定義の出力数", shown, "", "最大1万件。件数の集計は全名前を対象に継続。"
    AddSummary "名前定義の省略件数", problems - shown, "自動", "上限を超えた問題行は詳細のみ省略。", 0
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    RecordError "名前定義", "名前一覧の走査・出力", number, description
    AddSummary "名前定義の調査", NG, "要確認", "既に出力した問題行は残しています。集計は未完了です。"
End Sub

' 名前の各プロパティを独立取得し、一部の失敗でほかの判定を捨てない。
Private Function NameProperty(ByVal nm As Excel.Name, ByVal propertyId As Long, _
    ByVal context As String, ByRef succeeded As Boolean) As Variant
    Dim label As String
    succeeded = False
    On Error GoTo Failed
    Select Case propertyId
        Case 1
            label = "名前"
            NameProperty = nm.Name
        Case 2
            label = "参照先"
            NameProperty = nm.RefersTo
        Case 3
            label = "表示状態"
            NameProperty = nm.Visible
        Case Else: Err.Raise 5, , "未知の名前プロパティです。"
    End Select
    succeeded = True
    Exit Function
Failed:
    RecordError context, label, Err.Number, Err.Description
    NameProperty = NG
End Function

' 1件でも取得失敗があれば、部分的な件数を確定値として返さない。
Private Function CompleteCount(ByVal count As Long, ByVal failed As Boolean) As Variant
    If failed Then CompleteCount = NG Else CompleteCount = count
End Function

' 内部名はスコープ名と引用符を取り除いて大文字小文字を区別せず判定する。
Private Function IsInternalName(ByVal nameText As String) As Boolean
    Dim base As String
    base = LCase$(Mid$(nameText, InStrRev(nameText, "!") + 1))
    IsInternalName = (base Like "_xlfn.*" Or base Like "_xlpm.*" Or base Like "_xleta.*" Or _
        base Like "_xlnm.*" Or base = "_filterdatabase")
End Function

' 引用文字列・引用シート名・構造化参照内の文字を除いて参照切れを判定する。
Private Function HasBrokenReference(ByVal formula As String) As Boolean
    Dim i As Long, ch As String
    i = 1
    Do While i <= Len(formula)
        ch = Mid$(formula, i, 1)
        If ch = """" Or ch = "'" Then
            i = SkipQuoted(formula, i, ch)
        ElseIf ch = "[" Then
            i = SkipBrackets(formula, i)
        ElseIf UCase$(Mid$(formula, i, 5)) = "#REF!" Then
            HasBrokenReference = True
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

' 外部ブック修飾子と!の組み合わせを調べる。単なるTable[列名]を外部参照としない。
Private Function HasExternalReference(ByVal formula As String) As Boolean
    Dim i As Long, last As Long, j As Long, qualifier As String, ch As String
    i = 1
    Do While i <= Len(formula)
        ch = Mid$(formula, i, 1)
        If ch = """" Then
            i = SkipQuoted(formula, i, ch)
        ElseIf ch = "'" Then
            last = SkipQuoted(formula, i, ch)
            qualifier = Mid$(formula, i + 1, last - i - 2)
            If Mid$(formula, last, 1) = "!" Then
                If (InStr(qualifier, "[") > 0 And InStr(qualifier, "]") > 0) Or _
                    InStr(qualifier, "\") > 0 Or InStr(qualifier, "/") > 0 Then
                    HasExternalReference = True
                    Exit Function
                End If
            End If
            i = last
        ElseIf ch = "[" Then
            last = SkipBrackets(formula, i)
            j = last
            Do While j <= Len(formula)
                ch = Mid$(formula, j, 1)
                If ch = "!" Then
                    HasExternalReference = True
                    Exit Function
                End If
                If Not IsIdentifierChar(ch) And ch <> ":" Then Exit Do
                j = j + 1
            Loop
            i = last
        Else
            i = i + 1
        End If
    Loop
End Function

' ExcelリンクとOLEリンクをそれぞれ1リンク1行で一覧にする。
Private Sub AnalyzeLinks(ByVal wb As Excel.Workbook, ByVal linkType As XlLinkType, ByVal label As String)
    Dim links As Variant, i As Long, count As Long
    On Error GoTo Failed
    links = wb.LinkSources(linkType)
    If Not IsEmpty(links) Then
        count = UBound(links) - LBound(links) + 1
    End If
    AddSummary label & "数", count, "自動", "リンクの更新は呼び出しません。", TH_LINKS
    If count > 0 Then
        For i = LBound(links) To UBound(links)
            AddSummary label & " " & CStr(i - LBound(links) + 1), CStr(links(i)), "", ""
        Next i
    End If
    Exit Sub
Failed:
    RecordError "ブック", label & "一覧", Err.Number, Err.Description
    AddSummary label & "一覧", NG, "要確認", "取得済みのリンク行は残しています。"
End Sub

' 接続の一覧・起動時更新設定と、シートのQueryTable設定を概要に出力する。
Private Sub AnalyzeConnections(ByVal wb As Excel.Workbook)
    Dim connection As Excel.WorkbookConnection, ws As Excel.Worksheet
    Dim number As Long, count As Long
    On Error GoTo Failed
    count = wb.Connections.Count
    AddSummary "データ接続数", count, "自動", "接続文字列・認証情報は出力しません。", TH_CONNECTIONS
    For Each connection In wb.Connections
        number = number + 1
        AnalyzeConnection wb, connection, number
    Next connection
    For Each ws In wb.Worksheets
        AnalyzeQueryTables ws
    Next ws
    AnalyzePivotRefresh wb
    Exit Sub
Failed:
    RecordError "ブック", "データ接続一覧", Err.Number, Err.Description
    AddSummary "データ接続一覧", NG, "要確認", "出力済みの接続は残しています。"
End Sub

' 接続1件のプロパティの取得失敗を記録し、残りの接続調査は継続する。
Private Sub AnalyzeConnection(ByVal wb As Excel.Workbook, ByVal connection As Excel.WorkbookConnection, _
    ByVal index As Long)
    Dim nameText As String, kind As Long, refresh As Variant, label As String
    On Error GoTo Failed
    label = "データ接続 " & CStr(index)
    nameText = connection.Name
    kind = connection.Type
    AddSummary label, nameText, "", "種類=" & ConnectionTypeLabel(kind)
    Select Case kind
        Case xlConnectionTypeOLEDB
            refresh = connection.OLEDBConnection.RefreshOnFileOpen
        Case xlConnectionTypeODBC
            refresh = connection.ODBCConnection.RefreshOnFileOpen
        Case xlConnectionTypeDATAFEED
            refresh = connection.DataFeedConnection.RefreshOnFileOpen
        Case xlConnectionTypeTEXT, xlConnectionTypeWEB
            refresh = QueryRefreshForConnection(wb, nameText)
        Case xlConnectionTypeXMLMAP
            refresh = "対象外"
        Case xlConnectionTypeMODEL, xlConnectionTypeWORKSHEET, xlConnectionTypeNOSOURCE
            refresh = "対象外（内部接続。更新元の外部接続設定を参照）"
        Case Else
            RecordError nameText, "ファイルを開くときに更新する", vbObjectError + 2005, _
                "この接続種別の起動時更新設定を読み出す公開プロパティに対応していません。種類=" & CStr(kind)
            refresh = NG
    End Select
    If VarType(refresh) = vbBoolean Then
        AddSummary label & "：ファイルを開くときに更新する", YesNo(CBool(refresh)), _
            BooleanJudgment(CBool(refresh)), "RefreshOnFileOpen。設定の変更・更新はしません。"
    Else
        AddSummary label & "：ファイルを開くときに更新する", refresh, "", _
            "XMLマップ・内部接続は対象外。その他の取得失敗はエラー記録を参照。"
    End If
    Exit Sub
Failed:
    RecordError "データ接続 " & CStr(index), "接続名・種類・起動時更新設定", Err.Number, Err.Description
    AddSummary "データ接続 " & CStr(index), IIf(nameText = "", NG, nameText), "要確認", "一部の取得に失敗。"
    AddSummary "データ接続 " & CStr(index) & "：ファイルを開くときに更新する", NG, "要確認", "エラー記録を参照。"
End Sub

' Web・テキスト接続は、対応するQueryTableの起動時更新設定を調べる。
Private Function QueryRefreshForConnection(ByVal wb As Excel.Workbook, ByVal connectionName As String) As Variant
    Dim ws As Excel.Worksheet, query As Excel.QueryTable, found As Boolean, enabled As Boolean
    On Error GoTo Failed
    For Each ws In wb.Worksheets
        For Each query In ws.QueryTables
            If query.WorkbookConnection.Name = connectionName Then
                found = True
                If query.RefreshOnFileOpen Then enabled = True
            End If
        Next query
    Next ws
    If Not found Then
        Err.Raise vbObjectError + 2006, , "対応するQueryTableがなく、起動時更新設定を取得できませんでした。"
    End If
    QueryRefreshForConnection = enabled
    Exit Function
Failed:
    RecordError connectionName, "QueryTable経由の起動時更新設定", Err.Number, Err.Description
    QueryRefreshForConnection = NG
End Function

' ピボットキャッシュにも起動時更新設定があるため、接続一覧に併記する。
Private Sub AnalyzePivotRefresh(ByVal wb As Excel.Workbook)
    Dim cache As Excel.PivotCache, index As Long
    On Error GoTo Failed
    For Each cache In wb.PivotCaches
        index = index + 1
        AnalyzeOnePivotRefresh cache, index
    Next cache
    Exit Sub
Failed:
    RecordError "ブック", "ピボットキャッシュの起動時更新設定一覧", Err.Number, Err.Description
    AddSummary "ピボットキャッシュの起動時更新設定一覧", NG, "要確認", "エラー記録を参照。"
End Sub

' キャッシュごとのRefreshOnFileOpenを読み、Refreshを実行しない。
Private Sub AnalyzeOnePivotRefresh(ByVal cache As Excel.PivotCache, ByVal index As Long)
    Dim enabled As Boolean
    On Error GoTo Failed
    enabled = cache.RefreshOnFileOpen
    AddSummary "ピボットキャッシュ " & CStr(index) & "：ファイルを開くときに更新する", _
        YesNo(enabled), BooleanJudgment(enabled), "PivotCache.RefreshOnFileOpen。更新は呼び出しません。"
    Exit Sub
Failed:
    RecordError "ピボットキャッシュ " & CStr(index), "起動時更新設定", Err.Number, Err.Description
    AddSummary "ピボットキャッシュ " & CStr(index) & "：ファイルを開くときに更新する", _
        NG, "要確認", "エラー記録を参照。"
End Sub

' 接続に共通プロパティのないWeb・テキスト等もQueryTable側の設定を取得する。
Private Sub AnalyzeQueryTables(ByVal ws As Excel.Worksheet)
    Dim query As Excel.QueryTable, index As Long
    On Error GoTo Failed
    For Each query In ws.QueryTables
        index = index + 1
        AnalyzeQueryTable ws.Name, query, index
    Next query
    Exit Sub
Failed:
    RecordError ws.Name, "QueryTable一覧", Err.Number, Err.Description
    AddSummary "QueryTable一覧：" & ws.Name, NG, "要確認", "取得済みの行は残しています。"
End Sub

' QueryTableごとの名前、接続名、起動時更新設定を取得する。Refreshは呼び出さない。
Private Sub AnalyzeQueryTable(ByVal sheetName As String, ByVal query As Excel.QueryTable, ByVal index As Long)
    Dim label As String, nameText As String, refresh As Boolean, connectionName As Variant
    On Error GoTo Failed
    label = "QueryTable：" & sheetName & " / " & CStr(index)
    nameText = query.Name
    refresh = query.RefreshOnFileOpen
    connectionName = QueryConnectionName(query, sheetName)
    AddSummary label, nameText, "", "接続名=" & CStr(connectionName)
    AddSummary label & "：ファイルを開くときに更新する", YesNo(refresh), _
        BooleanJudgment(refresh), "QueryTable.RefreshOnFileOpen。更新は呼び出しません。"
    Exit Sub
Failed:
    RecordError sheetName, "QueryTable " & CStr(index) & " の起動時更新設定", Err.Number, Err.Description
    AddSummary "QueryTable：" & sheetName & " / " & CStr(index), NG, "要確認", "エラー記録を参照。"
End Sub

' QueryTableの対応接続名を取得する。取得失敗を空の接続名にしない。
Private Function QueryConnectionName(ByVal query As Excel.QueryTable, ByVal place As String) As Variant
    On Error GoTo Failed
    QueryConnectionName = query.WorkbookConnection.Name
    Exit Function
Failed:
    RecordError place, "QueryTableの接続名", Err.Number, Err.Description
    QueryConnectionName = NG
End Function

' 接続種別を既知の数値から表示名にする。未使用APIへの依存を増やさない。
Private Function ConnectionTypeLabel(ByVal kind As Long) As String
    Select Case kind
        Case 1: ConnectionTypeLabel = "OLE DB"
        Case 2: ConnectionTypeLabel = "ODBC"
        Case 3: ConnectionTypeLabel = "XMLマップ"
        Case 4: ConnectionTypeLabel = "テキスト"
        Case 5: ConnectionTypeLabel = "Web"
        Case 6: ConnectionTypeLabel = "データフィード"
        Case 7: ConnectionTypeLabel = "データモデル"
        Case 8: ConnectionTypeLabel = "ワークシート"
        Case 9: ConnectionTypeLabel = "ソースなし"
        Case Else: ConnectionTypeLabel = "その他(" & CStr(kind) & ")"
    End Select
End Function

' 起動時更新の有無の表示を統一する。
Private Function YesNo(ByVal value As Boolean) As String
    If value Then YesNo = "あり" Else YesNo = "なし"
End Function

' 自動更新が有効な接続は確認対象として表示する。
Private Function BooleanJudgment(ByVal value As Boolean) As String
    If value Then BooleanJudgment = "要確認"
End Function

' 対象を保存せず閉じる。開く処理中に中断した場合のブックも回収する。
Private Function CloseTarget() As Boolean
    Dim wb As Excel.Workbook, i As Long
    On Error GoTo Failed
    If mExcel Is Nothing Then
        CloseTarget = True
        Exit Function
    End If
    ' Workbooks.Openがエラーを返しても、途中まで開いたブックが残ることがある。
    ' このインスタンスで作った結果以外をすべて保存せず閉じる。
    For i = mExcel.Workbooks.Count To 1 Step -1
        Set wb = mExcel.Workbooks(i)
        If Not wb Is mReport Then wb.Close SaveChanges:=False
    Next i
    For i = mExcel.ProtectedViewWindows.Count To 1 Step -1
        mExcel.ProtectedViewWindows(i).Close
    Next i
    Set mTarget = Nothing
    CloseTarget = True
    Exit Function
Failed:
    RecordError "後始末", "調査対象を保存せず閉じる", Err.Number, Err.Description
    CloseTarget = False
End Function

' 設定を1つずつ復元する。AutomationSecurityは元のLowではなくByUIにする。
Private Sub RestoreSetting(ByRef old As ExcelSettings, ByVal index As Long)
    Dim label As String
    On Error GoTo Failed
    Select Case index
        Case 1
            label = "AutomationSecurity"
            mExcel.AutomationSecurity = msoAutomationSecurityByUI
        Case 2
            label = "Calculation"
            mExcel.Calculation = old.Calculation
        Case 3
            label = "EnableEvents"
            mExcel.EnableEvents = old.EnableEvents
        Case 4
            label = "DisplayAlerts"
            mExcel.DisplayAlerts = old.DisplayAlerts
        Case 5
            label = "ScreenUpdating"
            mExcel.ScreenUpdating = old.ScreenUpdating
        Case 6
            label = "StatusBar"
            mExcel.StatusBar = old.StatusBar
    End Select
    Exit Sub
Failed:
    RecordError "後始末", label & "の復元", Err.Number, Err.Description
End Sub

' 終了状態と最終エラー件数を反映し、取得済みの行は消さずに体裁を整える。
Private Sub FinalizeReport(ByVal aborted As Boolean)
    Dim ws As Excel.Worksheet
    On Error GoTo Failed
    If mReport Is Nothing Then Exit Sub
    If Not mSummary Is Nothing Then
        WriteRow mSummary, 3, Array("調査状態", IIf(aborted, "中断", "完了"), _
            IIf(aborted, "要確認", ""), "取得失敗の有無はエラー記録を参照。")
        mErrorCountRow = mSummaryRow
        AddSummary "エラー記録件数", mErrorCount, "自動", "取得失敗・中断・後始末の失敗を含みます。", 0
        mLogFailuresRow = mSummaryRow
        AddSummary "エラー記録の出力失敗件数", mLogFailures, "自動", "Excelが応答しない場合等は記録できません。", 0
    End If
    For Each ws In mReport.Worksheets
        FormatReportSheet ws
    Next ws
    Exit Sub
Failed:
    RecordError "出力ブック", "終了状態の反映", Err.Number, Err.Description
End Sub

' 書式設定や表示の段階で増えたエラーも、概要の最終件数に反映する。
Private Sub UpdateErrorCounts()
    On Error GoTo Failed
    If mSummary Is Nothing Then Exit Sub
    If mErrorCountRow > 0 Then
        WriteRow mSummary, mErrorCountRow, Array("エラー記録件数", mErrorCount, _
            JudgeNumber(mErrorCount, 0), "取得失敗・中断・後始末の失敗を含みます。")
    End If
    If mLogFailuresRow > 0 Then
        WriteRow mSummary, mLogFailuresRow, Array("エラー記録の出力失敗件数", mLogFailures, _
            JudgeNumber(mLogFailures, 0), "Excelが応答しない場合等は記録できません。")
    End If
    Exit Sub
Failed:
    RecordError "出力ブック", "最終エラー件数の反映", Err.Number, Err.Description
End Sub

' 巨大な参照文字列で列幅が広がらないよう、列幅の上限を設ける。
Private Sub FormatReportSheet(ByVal ws As Excel.Worksheet)
    Dim column As Excel.Range
    On Error GoTo Failed
    With ws.UsedRange
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(221, 235, 247)
        .Columns.AutoFit
        For Each column In .Columns
            If column.ColumnWidth > 60 Then column.ColumnWidth = 60
        Next column
        If .Rows.Count > 1 Then .AutoFilter
    End With
    Exit Sub
Failed:
    RecordError ws.Name, "結果の書式設定", Err.Number, Err.Description
End Sub

' 結果を開いたまま別インスタンスをユーザーへ渡す。エラー時もまず表示を試みる。
Private Function ShowResults() As Boolean
    On Error GoTo Failed
    If mExcel Is Nothing Then Exit Function
    mExcel.Visible = True
    If Not mReport Is Nothing Then
        mReport.Activate
        If Not mSummary Is Nothing Then mSummary.Activate
    End If
    mExcel.UserControl = True
    ShowResults = True
    Exit Function
Failed:
    RecordError "後始末", "別Excelの表示", Err.Number, Err.Description
    ShowResults = EnsureVisible()
End Function

' アクティブ化などが失敗しても、インスタンスだけは表示して非表示の残存を防ぐ。
Private Function EnsureVisible() As Boolean
    On Error GoTo Failed
    If mExcel Is Nothing Then Exit Function
    mExcel.Visible = True
    mExcel.UserControl = True
    EnsureVisible = True
    Exit Function
Failed:
    RecordError "後始末", "別Excelの再表示", Err.Number, Err.Description
    EnsureVisible = False
End Function

' 表示もできない異常時は、対象を保存せず閉じて専用インスタンスを終了する。
Private Sub QuitOwnedExcel()
    On Error GoTo Failed
    If mExcel Is Nothing Then Exit Sub
    mExcel.DisplayAlerts = False
    mExcel.Quit
    Exit Sub
Failed:
    RecordError "後始末", "別Excelの終了", Err.Number, Err.Description
    MsgBox "別Excelの表示・終了に失敗しました。タスクマネージャーで確認してください。" & _
        vbCrLf & mLastError, vbCritical
End Sub

' ファイルやExcel設定を変更せず、字句解析・判定の回帰テストだけを実行する。
Public Sub TestHeavyBookParser()
    Dim cases As Variant, entry As Variant, totals(0 To 8) As Double
    Dim i As Long, testCount As Long, mask As Long, expected As Long, quote As String
    On Error GoTo Failed
    quote = Chr$(34)
    cases = Array(Array("=INDIRECT(" & quote & "A1" & quote & ")", 1), _
        Array("=OFFSET(A1,0,1)+TODAY()+NOW()+RAND()+RANDBETWEEN(1,2)+CELL()+INFO()", 254), _
        Array("=TODAY()+TODAY()+NOW()+NOW()", 12), _
        Array("=" & quote & "NOW()" & quote, 0), _
        Array("=" & quote & "a" & quote & quote & "OFFSET()" & quote & "+RAND()", 16), _
        Array("='OFFSET('!A1+NOW()", 8), _
        Array("='a''NOW()'!A1", 0), _
        Array("=SUM(Table1[[#Headers],[NOW()]])", 0), _
        Array("=Table1[OFFSET(]+1", 0), _
        Array("=MYOFFSET(A1)+MYTODAY()+日本語NOW()", 0), _
        Array("=_xlfn._xlws.NOW()+today()", 12), _
        Array("=IF(TRUE,NOW(),RAND())", 24), _
        Array("=RANDARRAY(3)+RANDOM()+INFO(1)", 128))
    For Each entry In cases
        For i = 0 To 8: totals(i) = 0: Next i
        If HasVolatileName(UCase$(CStr(entry(0)))) Then CountFormulaText UCase$(CStr(entry(0))), totals
        mask = CLng(entry(1))
        expected = 0
        If mask <> 0 Then expected = 1
        AssertCondition totals(0) = expected, "重複除外: " & CStr(entry(0))
        For i = 1 To 8
            expected = 0
            If (mask And CLng(2 ^ (i - 1))) <> 0 Then expected = 1
            AssertCondition totals(i) = expected, "関数別: " & CStr(i) & " / " & CStr(entry(0))
        Next i
        testCount = testCount + 1
    Next entry
    cases = Array(Array("=SUM(Table1[Amount])", False), _
        Array("=SUM(Table1[Amount])+Sheet1!A1", False), _
        Array("=[Book.xlsx]Sheet1!A1", True), _
        Array("='C:\Data\[Book.xlsx]Sheet 1'!A1", True), _
        Array("='https://example.invalid/[Book.xlsx]Sheet1'!A1", True), _
        Array("=" & quote & "[Book.xlsx]Sheet1!A1" & quote, False), _
        Array("='Sheet 1'!A1", False))
    For Each entry In cases
        AssertCondition HasExternalReference(CStr(entry(0))) = CBool(entry(1)), _
            "外部参照: " & CStr(entry(0))
        testCount = testCount + 1
    Next entry
    AssertCondition HasBrokenReference("=#REF!+1"), "参照切れ"
    AssertCondition Not HasBrokenReference("=" & quote & "#REF!" & quote), "参照切れの文字列除外"
    AssertCondition Not HasBrokenReference("='名前#REF!'!A1"), "参照切れのシート名除外"
    AssertCondition Not HasBrokenReference("=Table1[#REF!]"), "参照切れの列名除外"
    AssertCondition IsInternalName("'a!b'!_XLFN.TEST"), "内部名のスコープ・大文字"
    AssertCondition Not IsInternalName("_xlfnUserName"), "通常名の除外禁止"
    AssertCondition IsNoCellsMessage("該当するセルが見つかりません。"), "日本語の該当なし"
    AssertCondition IsNoCellsMessage("No cells were found."), "英語の該当なし"
    AssertCondition Not IsNoCellsMessage("Application-defined or object-defined error"), "別の1004"
    AssertCondition JudgeNumber(TH_STYLES, TH_STYLES) = "", "しきい値ちょうど"
    AssertCondition JudgeNumber(TH_STYLES + 1, TH_STYLES) = "要確認", "しきい値超過"
    AssertCondition JudgeNumber(NG, TH_STYLES) = "要確認", "取得失敗の判定"
    AssertCondition CompleteCount(5, True) = NG, "部分集計は確定しない"
    MsgBox "字句解析・判定のテスト " & CStr(testCount + 13) & " ケースが成功しました。", vbInformation
    Exit Sub
Failed:
    MsgBox "字句解析・判定のテストに失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' 回帰テストの不一致を通常のVBAエラーとして報告する。
Private Sub AssertCondition(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 2100, "TestHeavyBookParser", label
End Sub
