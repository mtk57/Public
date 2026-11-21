Attribute VB_Name = "GenerateHeavyBookTests"
Option Explicit

' HeavyBookCheck用のテストファイル7個を、選択したローカルフォルダに作成する。
' Windows版Microsoft 365。追加参照不要。納品時の文字コードはCP932、改行はCRLF。

' 再計算負荷の調整用。生成時には計算せず、調査マクロで計測する。
Private Const CALC_FORMULA_COUNT As Long = 100001
Private Const SUMPRODUCT_FORMULA_COUNT As Long = 12000
Private Const SUMPRODUCT_SOURCE_ROWS As Long = 30000

' 件数の初期値。各項目を1枚のシートに集める。
Private Const HEAVY_CUSTOM_STYLES As Long = 2100
Private Const HEAVY_NORMAL_NAMES As Long = 1100
Private Const HEAVY_CF As Long = 110
Private Const HEAVY_SHAPES As Long = 510
Private Const HEAVY_NOTES As Long = 1010
Private Const HEAVY_THREADS As Long = 3
Private Const HEAVY_PIVOTS As Long = 11
Private Const HEAVY_CACHES As Long = 7
Private Const HEAVY_CHARTS As Long = 101
Private Const HEAVY_TABLES As Long = 101
Private Const HEAVY_LINKS As Long = 1010
Private Const HEAVY_VALIDATIONS As Long = 100001
Private Const EDGE_STYLES As Long = 2000
Private Const EDGE_NAMES As Long = 1000
Private Const EDGE_CF As Long = 100
Private Const EDGE_SHAPES As Long = 500
Private Const EDGE_NOTES As Long = 1000
Private Const EDGE_PIVOTS As Long = 10
Private Const EDGE_CACHES As Long = 6
Private Const EDGE_CHARTS As Long = 100
Private Const EDGE_TABLES As Long = 100
Private Const EDGE_LINKS As Long = 1000
Private Const EDGE_VALIDATIONS As Long = 100000
Private Const EDGE_FORMULAS As Long = 100000
Private Const FAILED As String = "取得失敗"
Private Const UNSUPPORTED As String = "非対応"
Private Const FILE_COLLISION As Long = vbObjectError + 4301

' ネットワークドライブを出力先にしないため、Windows標準APIで種別を確認する。
Private Declare PtrSafe Function GetDriveTypeW Lib "kernel32" (ByVal rootPath As LongPtr) As Long

Private Type SavedSettings
    Calculation As XlCalculation
    CalculateBeforeSave As Boolean
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    StatusBar As Variant
    AutomationSecurity As Long
End Type

Private gExcel As Excel.Application
Private gExpected As Excel.Workbook
Private gValues As Collection
Private gDetails As Collection
Private gFailures As Collection
Private gChecks As Collection
Private gFolder As String
Private gCollision As Boolean
Private gRunning As Boolean
Private gFile As String

' 出力先を選び、重い計算を実行せずにファイル一式と再取得した期待値を作成する。
Public Sub GenerateHeavyBookTestFiles()
    Dim dialog As FileDialog, old As SavedSettings, settingsSaved As Boolean
    Dim files As Variant, item As Variant, i As Long, aborted As Boolean
    Dim number As Long, description As String, message As String, reportSaved As Boolean, closed As Boolean
    If gRunning Then
        MsgBox "テストファイルの生成は実行中です。", vbExclamation
        Exit Sub
    End If
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "テストファイルを作成するローカルフォルダを選択"
    dialog.AllowMultiSelect = False
    If dialog.Show <> -1 Then Exit Sub
    gFolder = dialog.SelectedItems(1)
    If Right$(gFolder, 1) <> "\" Then gFolder = gFolder & "\"
    Set gValues = New Collection
    Set gDetails = New Collection
    Set gFailures = New Collection
    Set gChecks = New Collection
    gCollision = False
    gRunning = True
    On Error GoTo Failed
    ValidateConstants
    ValidateLocalFolder
    files = Array("test_clean.xlsx", "test_heavy.xlsx", "test_edge.xlsx", "test_legacy.xls", _
        "link_target.xlsx", "data.csv", "expected.xlsx")
    ' 1ファイルも作る前に、7ファイル全部を確認する。
    For Each item In files
        AssertNewPath gFolder & CStr(item)
    Next item

    Set gExcel = New Excel.Application
    Set gExpected = gExcel.Workbooks.Add(xlWBATWorksheet)
    SaveSettings old
    settingsSaved = True
    gExcel.ScreenUpdating = False
    gExcel.EnableEvents = False
    gExcel.DisplayAlerts = False
    gExcel.AutomationSecurity = msoAutomationSecurityForceDisable
    gExcel.Calculation = xlCalculationManual
    ' 手動計算でも保存直前に重い式を計算しない。
    gExcel.CalculateBeforeSave = False

    CreateCSV
    If gCollision Then GoTo Finish
    CreateLinkTarget
    If gCollision Then GoTo Finish
    For Each item In Array("test_clean.xlsx", "test_heavy.xlsx", "test_edge.xlsx", "test_legacy.xls")
        gExcel.StatusBar = "生成中: " & CStr(item)
        CreateTestBook CStr(item)
        If gCollision Then GoTo Finish
    Next item
    ' 期待値は保存後に開き直したブックから取得する。
    For Each item In Array("test_clean.xlsx", "test_heavy.xlsx", "test_edge.xlsx", _
        "test_legacy.xls", "link_target.xlsx")
        gExcel.StatusBar = "保存済みファイルを再集計: " & CStr(item)
        InspectSavedBook CStr(item)
    Next item
    InspectCSV
    BuildChecks
    GoTo Finish
Failed:
    number = Err.Number
    description = Err.Description
    aborted = True
    AddFailure "全体", "生成を中断", number, description
    Resume Finish
Finish:
    closed = CloseOtherBooks()
    If settingsSaved Then
        For i = 1 To 7
            If Not closed And (i = 2 Or i = 3) Then
                AddFailure "後始末", "計算設定の復元", vbObjectError + 4302, _
                    "対象を閉じられないため、重い式の未許可計算を防ぐ設定を維持しました。"
            Else
                RestoreOneSetting old, i
            End If
        Next i
    End If
    If Not gExpected Is Nothing And Not gCollision Then reportSaved = SaveExpected()
    CloseAndQuit
    message = "テストファイルの生成が完了しました。"
    If aborted Or gCollision Then message = "テストファイルの生成を中断しました。既存ファイルは上書きしていません。"
    message = message & vbCrLf & gFolder
    If Not reportSaved Then message = message & vbCrLf & "expected.xlsxは保存できていません。"
    If gFailures.Count > 0 Then
        message = message & vbCrLf & "失敗・未対応項目: " & CStr(gFailures.Count) & " 件"
        ShowFailureList message, reportSaved
    Else
        MsgBox message & vbCrLf & "expected.xlsxと調査結果を照合してください。", vbInformation
    End If
    Set gExpected = Nothing
    Set gExcel = Nothing
    gRunning = False
End Sub

' 負荷の定数がExcelの行数を超えたり、重い式が総数を超えたりしないよう確認する。
Private Sub ValidateConstants()
    If CALC_FORMULA_COUNT < 1 Or CALC_FORMULA_COUNT > 1048576 Then Err.Raise 5, , "CALC_FORMULA_COUNTが範囲外です。"
    If SUMPRODUCT_FORMULA_COUNT < 1 Or SUMPRODUCT_FORMULA_COUNT > CALC_FORMULA_COUNT Then _
        Err.Raise 5, , "SUMPRODUCT_FORMULA_COUNTは1以上、CALC_FORMULA_COUNT以下にしてください。"
    If SUMPRODUCT_SOURCE_ROWS < 1 Or SUMPRODUCT_SOURCE_ROWS > 1048576 Then _
        Err.Raise 5, , "SUMPRODUCT_SOURCE_ROWSが範囲外です。"
End Sub

' UNC・URL・ネットワークドライブを除外し、ローカルディスクだけを許可する。
Private Sub ValidateLocalFolder()
    Dim root As String, kind As Long
    If Len(gFolder) < 3 Then Err.Raise 5, , "ローカルフォルダを指定してください。"
    If Mid$(gFolder, 2, 2) <> ":\" Then Err.Raise 5, , "UNCやURLは使用できません。"
    root = Left$(gFolder, 3)
    kind = GetDriveTypeW(StrPtr(root))
    If kind <> 2 And kind <> 3 And kind <> 6 Then Err.Raise 5, , "ネットワークドライブ等は使用できません。"
End Sub

' 隠しファイル・同名フォルダも含め、既存の出力パスを拒否する。
Private Sub AssertNewPath(ByVal path As String)
    If PathExists(path) Then
        gCollision = True
        Err.Raise FILE_COLLISION, , "同名ファイルまたはフォルダが存在します: " & path
    End If
End Sub

' 指定パスの存在を確認する。確認自体の失敗は上位のハンドラーへ返す。
Private Function PathExists(ByVal path As String) As Boolean
    PathExists = (Len(Dir$(path, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory Or vbArchive)) > 0)
End Function

' 専用Excelの変更前の設定を保存する。
Private Sub SaveSettings(ByRef old As SavedSettings)
    With gExcel
        old.Calculation = .Calculation
        old.CalculateBeforeSave = .CalculateBeforeSave
        old.ScreenUpdating = .ScreenUpdating
        old.EnableEvents = .EnableEvents
        old.DisplayAlerts = .DisplayAlerts
        old.StatusBar = .StatusBar
        old.AutomationSecurity = .AutomationSecurity
    End With
End Sub

' 設定の復元は1項目失敗しても続ける。残っているのは期待値の値だけのブック。
Private Sub RestoreOneSetting(ByRef old As SavedSettings, ByVal index As Long)
    On Error GoTo Failed
    Select Case index
        Case 1: gExcel.AutomationSecurity = old.AutomationSecurity
        Case 2: gExcel.Calculation = old.Calculation
        Case 3: gExcel.CalculateBeforeSave = old.CalculateBeforeSave
        Case 4: gExcel.EnableEvents = old.EnableEvents
        Case 5: gExcel.DisplayAlerts = old.DisplayAlerts
        Case 6: gExcel.ScreenUpdating = old.ScreenUpdating
        Case 7: gExcel.StatusBar = old.StatusBar
    End Select
    Exit Sub
Failed:
    AddFailure "後始末", "Excel設定の復元 " & CStr(index), Err.Number, Err.Description
End Sub

' CSVは同じフォルダにASCIIで作る。通信や外部コマンドは使わない。
Private Sub CreateCSV()
    Dim handle As Integer, opened As Boolean, i As Long, number As Long, description As String
    On Error GoTo Failed
    AssertNewPath gFolder & "data.csv"
    handle = FreeFile
    Open gFolder & "data.csv" For Output Access Write Lock Read Write As #handle
    opened = True
    Print #handle, "Category,Amount"
    For i = 1 To 100
        Print #handle, "Group" & CStr(i Mod 4) & "," & CStr(i)
    Next i
    Close #handle
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    If opened Then CloseFileHandle handle
    AddFailure "data.csv", "CSV作成", number, description
End Sub

' 開いたCSVハンドルを、元のエラー情報を壊さず閉じる。
Private Sub CloseFileHandle(ByVal handle As Integer)
    On Error GoTo Failed
    Close #handle
    Exit Sub
Failed:
    AddFailure "data.csv", "CSVを閉じる", Err.Number, Err.Description
End Sub

' ローカル外部参照の相手を先に作る。
Private Sub CreateLinkTarget()
    Dim wb As Excel.Workbook
    On Error GoTo Failed
    Set wb = gExcel.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Data"
    wb.Worksheets(1).Range("A1").Value2 = 123
    SaveNewBook wb, "link_target.xlsx", xlOpenXMLWorkbook
    CloseOneBook wb
    Exit Sub
Failed:
    AddFailure "link_target.xlsx", "外部リンク先の作成", Err.Number, Err.Description
    If Not wb Is Nothing Then CloseOneBook wb
End Sub

' ファイル種別に応じた項目を作る。各機能の失敗はTryFeature内で記録して継続する。
Private Sub CreateTestBook(ByVal fileName As String)
    Dim wb As Excel.Workbook, feature As Long
    On Error GoTo Failed
    gFile = fileName
    Set wb = gExcel.Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "基本"
    wb.Worksheets(1).Range("A1:B3").Value2 = Array2D(Array(Array(1, 10), Array(2, 20), Array(3, 30)))
    wb.Worksheets(1).Range("C1").Formula = "=SUM(A1:B3)"
    Select Case fileName
        Case "test_heavy.xlsx"
            For feature = 1 To 16
                TryFeature wb, feature, False
            Next feature
        Case "test_edge.xlsx"
            For feature = 1 To 11
                TryFeature wb, feature, True
            Next feature
            TryFeature wb, 13, True
            TryFeature wb, 17, True
            ' スタイルと名前の総数を最後にちょうどに合わせる。
            TryFeature wb, 14, True
            TryFeature wb, 15, True
            ProtectEdgeSheet wb
    End Select
    If fileName = "test_legacy.xls" Then
        SaveNewBook wb, fileName, xlExcel8
    Else
        SaveNewBook wb, fileName, xlOpenXMLWorkbook
    End If
    CloseOneBook wb
    Exit Sub
Failed:
    AddFailure fileName, "ブックの作成・保存", Err.Number, Err.Description
    If Not wb Is Nothing Then CloseOneBook wb
End Sub

' 機能ごとに独立して生成し、特定APIが使えない場合も次の項目を作る。
Private Sub TryFeature(ByVal wb As Excel.Workbook, ByVal feature As Long, ByVal edge As Boolean)
    On Error GoTo Failed
    gExcel.StatusBar = "生成中: " & gFile & " / " & FeatureLabel(feature)
    Select Case feature
        Case 1: MakeExtraRows wb, edge
        Case 2: MakeExtraColumns wb, edge
        Case 3
            If Not edge Then MakeVolatile wb
        Case 4: MakeConditionalFormats wb, edge
        Case 5: MakeShapes wb, edge
        Case 6: MakeNotes wb, edge
        Case 7: MakePivots wb, edge
        Case 8: MakeCharts wb, edge
        Case 9: MakeTables wb, edge
        Case 10: MakeHyperlinks wb, edge
        Case 11: MakeValidation wb, edge
        Case 12: MakeCSVConnection wb
        Case 13: MakeCalculation wb, edge
        Case 14: MakeStyles wb, edge
        Case 15: MakeNames wb, edge
        Case 16: MakeThreadedComments wb
        Case 17: MakeEdgeCases wb
    End Select
    Exit Sub
Failed:
    AddFailure gFile, FeatureLabel(feature), Err.Number, Err.Description
End Sub

' 生成失敗の一覧に用いる項目名を返す。
Private Function FeatureLabel(ByVal feature As Long) As String
    Dim labels As Variant
    labels = Array("", "余分な行", "余分な列", "揮発性関数", "条件付き書式", "図形", "メモ", _
        "ピボット・キャッシュ", "グラフ", "テーブル", "ハイパーリンク", "入力規則", "CSV接続", _
        "計算式", "スタイル", "名前定義", "スレッドコメント", "誤検出・シート種別")
    FeatureLabel = CStr(labels(feature))
End Function

' 新しいワークシートを追加する。
Private Function AddSheet(ByVal wb As Excel.Workbook, ByVal nameText As String) As Excel.Worksheet
    Dim ws As Excel.Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nameText
    Set AddSheet = ws
End Function

' 定数だけのA1:D100を配列で一括設定する。
Private Sub FillData100(ByVal ws As Excel.Worksheet)
    Dim data(1 To 100, 1 To 4) As Variant, r As Long, c As Long
    For r = 1 To 100
        For c = 1 To 4
            data(r, c) = r * 10 + c
        Next c
    Next r
    ws.Range("A1:D100").Value2 = data
End Sub

' 実データはA1:D100だけにし、遠い行の書式で使用範囲を伸ばす。
Private Sub MakeExtraRows(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet
    Set ws = AddSheet(wb, "余分な行")
    FillData100 ws
    If edge Then
        ws.Range("Z1100").Interior.Color = RGB(255, 255, 0)
    Else
        ws.Range("Z60000").Interior.Color = RGB(255, 255, 0)
    End If
End Sub

' D列から100列または101列だけ余分な列を作る。
Private Sub MakeExtraColumns(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet
    Set ws = AddSheet(wb, "余分な列")
    FillData100 ws
    If edge Then
        ws.Range("CZ100").Interior.Color = RGB(255, 255, 0)
    Else
        ws.Range("DA100").Interior.Color = RGB(255, 255, 0)
    End If
End Sub

' 各関数を1,3,...15セル、さらに複数関数と同一関数の重複を3セルに配置する。
Private Sub MakeVolatile(ByVal wb As Excel.Workbook)
    Dim ws As Excel.Worksheet, formulas As Variant, i As Long, j As Long, r As Long
    Set ws = AddSheet(wb, "揮発性関数")
    ws.Range("D1").Value2 = 1
    formulas = VolatileTemplates()
    r = 1
    For i = 0 To 7
        For j = 1 To i * 2 + 1
            ws.Cells(r, 1).Formula = formulas(i)
            r = r + 1
        Next j
    Next i
    For i = 8 To 10
        ws.Cells(r, 1).Formula = formulas(i)
        r = r + 1
    Next i
    ws.Range("B1").Formula = "='" & Replace(gFolder, "'", "''") & "[link_target.xlsx]Data'!$A$1"
End Sub

' 生成・期待値照合に用いる既知の式。任意数式の字句解析は再利用しない。
Private Function VolatileTemplates() As Variant
    VolatileTemplates = Array("=INDIRECT(""D1"")", "=OFFSET($D$1,0,0)", "=TODAY()", "=NOW()", _
        "=RAND()", "=RANDBETWEEN(1,100)", "=CELL(""address"",$D$1)", "=INFO(""system"")", _
        "=TODAY()+NOW()", "=RAND()+RAND()+RANDBETWEEN(1,100)", _
        "=SUM(OFFSET($D$1,0,0),INDIRECT(""D1""))")
End Function

' 1枚のシートに独立した条件付き書式ルールを作る。
Private Sub MakeConditionalFormats(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, i As Long, n As Long, rule As Excel.FormatCondition
    Set ws = AddSheet(wb, "条件付き書式")
    n = HEAVY_CF
    If edge Then n = EDGE_CF
    ws.Range("A1").Resize(n, 1).Value2 = 1
    For i = 1 To n
        Set rule = ws.Cells(i, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=A" & CStr(i) & "=" & CStr(i))
        rule.Interior.Color = RGB(i Mod 255, 200, 100)
    Next i
End Sub

' 通常図形を集約し、heavyだけ非表示5個・1pt未満3個を作る。
Private Sub MakeShapes(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, shp As Excel.Shape, i As Long, n As Long
    Set ws = AddSheet(wb, "図形")
    n = HEAVY_SHAPES
    If edge Then n = EDGE_SHAPES
    For i = 1 To n
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, (i Mod 25) * 12, (i \ 25) * 12, 8, 8)
        shp.Name = "TestShape_" & Format$(i, "0000")
        If Not edge Then
            If i <= 5 Then shp.Visible = msoFalse
            If i >= 6 And i <= 8 Then
                shp.LockAspectRatio = msoFalse
                shp.Width = 0.5
            End If
        End If
    Next i
End Sub

' 従来のコメント（メモ）を1枚に集約する。
Private Sub MakeNotes(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, i As Long, n As Long
    Set ws = AddSheet(wb, "コメント")
    n = HEAVY_NOTES
    If edge Then n = EDGE_NOTES
    For i = 1 To n
        ws.Cells(i, 1).AddComment "テスト用メモ " & CStr(i)
    Next i
End Sub

' スレッドコメントは同じコメントシートの別セルに追加する。未対応ならこの項目だけ失敗扱い。
Private Sub MakeThreadedComments(ByVal wb As Excel.Workbook)
    Dim cell As Object, i As Long
    For i = 1 To HEAVY_THREADS
        Set cell = wb.Worksheets("コメント").Cells(i, 2)
        cell.AddCommentThreaded "テスト用スレッド " & CStr(i)
    Next i
End Sub

' 異なるローカル範囲を元にキャッシュを作り、全ピボットを1枚へ集約する。
Private Sub MakePivots(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, source As Excel.Worksheet, caches As Collection
    Dim cache As Excel.PivotCache, pivot As Excel.PivotTable, data(1 To 101, 1 To 2) As Variant
    Dim i As Long, r As Long, cacheCount As Long, pivotCount As Long, start As Long
    Set ws = AddSheet(wb, "ピボット")
    Set source = AddSheet(wb, "ピボット元")
    Set caches = New Collection
    cacheCount = HEAVY_CACHES
    pivotCount = HEAVY_PIVOTS
    If edge Then
        cacheCount = EDGE_CACHES
        pivotCount = EDGE_PIVOTS
    End If
    data(1, 1) = "分類"
    data(1, 2) = "金額"
    For i = 1 To cacheCount
        start = 1 + (i - 1) * 103
        For r = 2 To 101
            data(r, 1) = "分類" & CStr(r Mod 3)
            data(r, 2) = r + i
        Next r
        source.Cells(start, 1).Resize(101, 2).Value2 = data
        Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:="'" & source.Name & "'!" & source.Cells(start, 1).Resize(101, 2).Address(ReferenceStyle:=xlR1C1))
        caches.Add cache
    Next i
    For i = 1 To pivotCount
        Set cache = caches(1 + (i - 1) Mod cacheCount)
        Set pivot = cache.CreatePivotTable(TableDestination:=ws.Cells(1 + (i - 1) * 15, 1), TableName:="TestPivot_" & CStr(i))
        pivot.PivotFields("分類").Orientation = xlRowField
        pivot.AddDataField pivot.PivotFields("金額"), "金額合計", xlSum
    Next i
End Sub

' 埋め込みグラフを1枚のシートに集約する。
Private Sub MakeCharts(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, chart As Excel.ChartObject, i As Long, n As Long
    Set ws = AddSheet(wb, "グラフ")
    ws.Range("A1:B3").Value2 = Array2D(Array(Array("分類", "値"), Array("A", 1), Array("B", 2)))
    n = HEAVY_CHARTS
    If edge Then n = EDGE_CHARTS
    For i = 1 To n
        Set chart = ws.ChartObjects.Add((i Mod 10) * 65, (i \ 10) * 45, 60, 40)
        chart.Chart.ChartType = xlColumnClustered
        chart.Chart.SetSourceData ws.Range("A1:B3")
    Next i
End Sub

' 互いに重ならない小さなテーブルを1枚に集約する。
Private Sub MakeTables(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, table As Excel.ListObject, i As Long, n As Long
    Set ws = AddSheet(wb, "テーブル")
    n = HEAVY_TABLES
    If edge Then n = EDGE_TABLES
    For i = 1 To n
        ws.Cells(1 + (i - 1) * 4, 1).Resize(2, 2).Value2 = Array2D(Array(Array("Key", "Amount"), Array(i, i * 2)))
        Set table = ws.ListObjects.Add(SourceType:=xlSrcRange, _
            Source:=ws.Cells(1 + (i - 1) * 4, 1).Resize(2, 2), XlListObjectHasHeaders:=xlYes)
        table.Name = "GeneratedTable_" & CStr(i)
    Next i
End Sub

' Webを参照せず、ブック内リンクだけを1枚に集約する。
Private Sub MakeHyperlinks(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, i As Long, n As Long
    Set ws = AddSheet(wb, "ハイパーリンク")
    n = HEAVY_LINKS
    If edge Then n = EDGE_LINKS
    For i = 1 To n
        ws.Hyperlinks.Add Anchor:=ws.Cells(i, 1), Address:="", SubAddress:="'基本'!A1", TextToDisplay:="内部リンク " & CStr(i)
    Next i
End Sub

' 入力規則を定数入りの範囲へ一括設定し、余分な使用範囲の警告を混ぜない。
Private Sub MakeValidation(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, n As Long
    Set ws = AddSheet(wb, "入力規則")
    n = HEAVY_VALIDATIONS
    If edge Then n = EDGE_VALIDATIONS
    ws.Range("A1").Resize(n, 1).Value2 = 1
    ws.Range("A1").Resize(n, 1).Validation.Add Type:=xlValidateWholeNumber, _
        AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="10"
End Sub

' 生成済みのローカルCSVだけを一度取り込み、起動時更新を有効にする。
Private Sub MakeCSVConnection(ByVal wb As Excel.Workbook)
    Dim ws As Excel.Worksheet, query As Excel.QueryTable
    If Not PathExists(gFolder & "data.csv") Then Err.Raise 53, , "data.csvの生成に失敗しています。"
    Set ws = AddSheet(wb, "CSV接続")
    Set query = ws.QueryTables.Add(Connection:="TEXT;" & gFolder & "data.csv", Destination:=ws.Range("A1"))
    query.TextFileParseType = xlDelimited
    query.TextFileCommaDelimiter = True
    query.BackgroundQuery = False
    query.Refresh BackgroundQuery:=False
    query.RefreshOnFileOpen = True
End Sub

' 総数の境界と、揮発性関数を使わないSUMPRODUCTの重い計算を作る。
Private Sub MakeCalculation(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, data() As Variant, i As Long
    Set ws = AddSheet(wb, "重い計算")
    If edge Then
        ws.Range("B1").Resize(EDGE_FORMULAS, 1).Formula = "=SUM(1,2,3)"
    Else
        ReDim data(1 To SUMPRODUCT_SOURCE_ROWS, 1 To 1)
        For i = 1 To SUMPRODUCT_SOURCE_ROWS
            data(i, 1) = (i Mod 100) / 100
        Next i
        ws.Range("A1").Resize(SUMPRODUCT_SOURCE_ROWS, 1).Value2 = data
        ws.Range("B1").Resize(CALC_FORMULA_COUNT, 1).Formula = "=SUM(1,2,3)"
        ws.Range("B1").Resize(SUMPRODUCT_FORMULA_COUNT, 1).FormulaR1C1 = _
            "=SUMPRODUCT((R1C1:R" & CStr(SUMPRODUCT_SOURCE_ROWS) & "C1+MOD(ROW(),17))*R1C1:R" & _
            CStr(SUMPRODUCT_SOURCE_ROWS) & "C1)"
    End If
End Sub

' 異なる見本セルに基づく名前付きスタイルを作り、実際に適用して保存時の保持を確実にする。
Private Sub MakeStyles(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim ws As Excel.Worksheet, sample As Excel.Range, style As Excel.Style
    Dim i As Long, n As Long, styleName As String
    Set ws = AddSheet(wb, "スタイル見本")
    ws.Range("A1:B1").Value2 = Array2D(Array(Array("スタイル名", "見本")))
    n = HEAVY_CUSTOM_STYLES
    If edge Then n = EDGE_STYLES - wb.Styles.Count
    If n < 0 Then Err.Raise 5, , "既存スタイル数が境界値を超えています。"
    For i = 1 To n
        styleName = "GeneratedStyle_" & Format$(i, "0000")
        Set sample = ws.Cells(i + 1, 2)
        sample.Value2 = i
        sample.Interior.Color = RGB(i Mod 256, (i \ 256) Mod 256, 128)
        ' Addの返り値は使わず、追加後のコレクションから名前で取り直す。
        wb.Styles.Add Name:=styleName, BasedOn:=sample
        Set style = wb.Styles.Item(styleName)
        If StrComp(style.Name, styleName, vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 4303, , "要求と異なるスタイル名が返されました: " & styleName & " / " & style.Name
        End If
        sample.Style = styleName
        ws.Cells(i + 1, 1).Value2 = styleName
    Next i
End Sub

' 通常名と問題名を生成する。edgeはMYOFFSET等を含む総数1000に合わせる。
Private Sub MakeNames(ByVal wb As Excel.Workbook, ByVal edge As Boolean)
    Dim i As Long, n As Long
    n = HEAVY_NORMAL_NAMES
    If edge Then n = EDGE_NAMES - wb.Names.Count
    If n < 0 Then Err.Raise 5, , "既存の名前定義数が境界値を超えています。"
    For i = 1 To n
        wb.Names.Add Name:="GeneratedName_" & Format$(i, "0000"), RefersTo:="=1"
    Next i
    If edge Then Exit Sub
    For i = 1 To 3
        wb.Names.Add Name:="BrokenReference_" & CStr(i), RefersTo:="=#REF!"
    Next i
    For i = 1 To 2
        wb.Names.Add Name:="ExternalReference_" & CStr(i), _
            RefersTo:="='" & Replace(gFolder, "'", "''") & "[link_target.xlsx]Data'!$A$1"
    Next i
    For i = 1 To 4
        wb.Names.Add Name:="HiddenName_" & CStr(i), RefersTo:="=1", Visible:=False
    Next i
    AddInternalHiddenName wb
End Sub

' 予約名を直接追加せず、AutoFilterにExcel自身の内部非表示名を作成させる。
Private Sub AddInternalHiddenName(ByVal wb As Excel.Workbook)
    Dim ws As Excel.Worksheet
    On Error GoTo Failed
    Set ws = AddSheet(wb, "内部名")
    ws.Range("A1:B3").Value2 = Array2D(Array(Array("Key", "Amount"), Array(1, 10), Array(2, 20)))
    ws.Range("A1:B3").AutoFilter
    ' _FilterDatabaseの存在と非表示状態は保存後のInspectNamesで確認する。
    Exit Sub
Failed:
    AddFailure gFile, "内部非表示名（_FilterDatabaseの自動作成）", Err.Number, Err.Description
End Sub

' シート種別と関数名に似た文字列・識別子を含む境界ケースを作る。
Private Sub MakeEdgeCases(ByVal wb As Excel.Workbook)
    Dim ws As Excel.Worksheet, named As Excel.Worksheet, table As Excel.ListObject, chart As Excel.Chart
    Set ws = AddSheet(wb, "空")
    Set ws = AddSheet(wb, "非表示")
    ws.Visible = xlSheetHidden
    Set ws = AddSheet(wb, "完全非表示")
    ws.Visible = xlSheetVeryHidden
    Set ws = AddSheet(wb, "保護")
    ws.Range("A1").Value2 = 1
    ' スタイル追加はブック内に保護シートがあると失敗するため、保護は作成の最後に行う。
    Set named = AddSheet(wb, "OFFSET(")
    named.Range("A1").Value2 = 10
    Set ws = AddSheet(wb, "誤検出")
    ws.Range("D1:E2").Value2 = Array2D(Array(Array("NOW()", "Amount"), Array(1, 2)))
    Set table = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("D1:E2"), XlListObjectHasHeaders:=xlYes)
    table.Name = "Table1"
    wb.Names.Add Name:="MYOFFSET", RefersTo:="=1"
    wb.Names.Add Name:="StructuredReference", RefersTo:="=SUM(Table1[Amount])"
    ws.Range("A1").Formula = "=""NOW()"""
    ws.Range("A2").Formula = "='OFFSET('!A1"
    ws.Range("A3").Formula = "=SUM(Table1[NOW()])"
    ws.Range("A4").Formula2 = "=RANDARRAY(1,1)"
    ws.Range("A5").Formula = "=MYOFFSET"
    ws.Range("A6").Formula = "=SUM(Table1[Amount])"
    Set chart = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))
    chart.Name = "グラフシート"
    chart.SetSourceData wb.Worksheets("基本").Range("A1:B3")
End Sub


' スタイル・名前をすべて作成した後、保存前に境界テスト用シートを保護する。
Private Sub ProtectEdgeSheet(ByVal wb As Excel.Workbook)
    On Error GoTo Failed
    wb.Worksheets("保護").Protect Password:="test"
    Exit Sub
Failed:
    AddFailure gFile, "保護シートの保護適用", Err.Number, Err.Description
End Sub

' 保存の直前にも存在を確認し、マクロを含まない形式を明示する。
Private Sub SaveNewBook(ByVal wb As Excel.Workbook, ByVal fileName As String, ByVal fileFormat As XlFileFormat)
    AssertNewPath gFolder & fileName
    wb.CheckCompatibility = False
    wb.SaveAs Filename:=gFolder & fileName, FileFormat:=fileFormat, CreateBackup:=False, AddToMru:=False
End Sub

' 入れ子のArrayを、Rangeへ渡せる2次元配列へ変換する。
Private Function Array2D(ByVal rows As Variant) As Variant
    Dim data() As Variant, r As Long, c As Long, row As Variant
    ReDim data(1 To UBound(rows) + 1, 1 To UBound(rows(0)) + 1)
    For r = 0 To UBound(rows)
        row = rows(r)
        For c = 0 To UBound(row)
            data(r + 1, c + 1) = row(c)
        Next c
    Next r
    Array2D = data
End Function

' 保存済みファイルを手動計算・リンク更新なし・読み取り専用で開き直す。
Private Sub InspectSavedBook(ByVal fileName As String)
    Dim wb As Excel.Workbook, ws As Excel.Worksheet, seconds As Double, started As Double
    Dim style As Excel.Style, custom As Long, sheet As Object
    On Error GoTo Failed
    gFile = fileName
    If Not PathExists(gFolder & fileName) Then Err.Raise 53, , "生成済みファイルがありません。作成記録を参照してください。"
    started = Timer
    Set wb = gExcel.Workbooks.Open(Filename:=gFolder & fileName, ReadOnly:=True, _
        UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    seconds = SecondsSince(started)
    ExpectValue fileName, "ブック", "ファイルサイズ(MB)", FileLen(gFolder & fileName) / 1048576#, -1, _
        "超過は目指さず実測のみ。調査側の判定基準は50MB超。"
    ExpectValue fileName, "ブック", "開く時間(秒)", seconds, -1, _
        "保存後の再オープン実測。キャッシュ・PC性能等で変わる。調査側は10秒超で要確認。"
    ExpectValue fileName, "ブック", "ファイル形式", wb.FileFormat, -1, _
        "ExcelのFileFormat実測。xlsは形式の判定が要確認。", IIf(wb.FileFormat = xlExcel8, "要確認", "")
    ExpectValue fileName, "ブック", "セルスタイル数", wb.Styles.Count, 2000
    custom = 0
    For Each style In wb.Styles
        If Not style.BuiltIn Then custom = custom + 1
    Next style
    ExpectValue fileName, "ブック", "うちユーザー定義スタイル数", custom
    InspectNames wb, fileName
    InspectLinks wb, fileName, xlExcelLinks, "Excel外部リンク"
    InspectLinks wb, fileName, xlOLELinks, "OLE外部リンク"
    ExpectValue fileName, "ブック", "ピボットキャッシュ数", wb.PivotCaches.Count, 6
    InspectConnections wb, fileName
    ExpectValue fileName, "ブック", "ワークシート数", wb.Worksheets.Count
    ExpectValue fileName, "ブック", "全シート数", wb.Sheets.Count, -1, "グラフシートはシート別調査の対象外。"
    For Each sheet In wb.Sheets
        AddDetail fileName, CStr(sheet.Name), "シート種別", TypeName(sheet), "", ""
    Next sheet
    For Each ws In wb.Worksheets
        InspectSheet ws, fileName
    Next ws
    ExpectValue fileName, "ブック", "再計算時間合計(秒)", "調査時に計測", -1, _
        "調査で許可した場合の専用Excel.CalculateFullの時間。生成中には計算しない。", "実測で判定"
    CloseOneBook wb
    Exit Sub
Failed:
    AddFailure fileName, "保存後の期待値取得", Err.Number, Err.Description
    ExpectValue fileName, "ブック", "期待値の再取得", FAILED, -1, "失敗するまでに取得した行は保持。"
    If Not wb Is Nothing Then CloseOneBook wb
End Sub

' シート単位でAPIから再集計し、シート単独のしきい値との比較を出力する。
Private Sub InspectSheet(ByVal ws As Excel.Worksheet, ByVal fileName As String)
    Dim last As Excel.Range, hit As Excel.Range, usedRow As Long, usedCol As Long
    Dim dataRow As Long, dataCol As Long, address As String, counts As Variant
    Dim shp As Excel.Shape, shapes As Long, hidden As Long, tiny As Long
    Dim total As Variant, threads As Variant, notes As Variant, i As Long
    Dim labels As Variant, label As String, usedNote As String
    On Error GoTo Failed
    label = "表示状態"
    ExpectValue fileName, ws.Name, label, VisibilityText(ws.Visible)
    label = "使用範囲・実データ"
    Set last = ExpectedLastCell(ws, usedNote)
    usedRow = last.Row
    usedCol = last.Column
    ExpectValue fileName, ws.Name, "使用範囲の最終セル(Ctrl+End)", last.Address(False, False, xlA1), -1, usedNote
    Set hit = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If Not hit Is Nothing Then dataRow = hit.Row
    Set hit = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If Not hit Is Nothing Then dataCol = hit.Column
    If dataRow = 0 And dataCol = 0 Then
        address = "(実データなし)"
        dataRow = 1
        dataCol = 1
    Else
        address = ws.Cells(dataRow, dataCol).Address(False, False, xlA1)
    End If
    ExpectValue fileName, ws.Name, "実データの最終セル", address
    ExpectValue fileName, ws.Name, "余分な行数", Nonnegative(usedRow - dataRow), 1000
    ExpectValue fileName, ws.Name, "余分な列数", Nonnegative(usedCol - dataCol), 100
    counts = InspectFormulas(ws, fileName)
    ExpectValue fileName, ws.Name, "数式セル数", counts(0), 100000
    ExpectValue fileName, ws.Name, "揮発性関数を含むセル数", counts(1), 0, _
        "生成用の既知の式と保存後の式を完全一致で照合。調査マクロの字句解析は呼ばない。"
    labels = Array("INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN", "CELL", "INFO")
    For i = 0 To 7
        ExpectValue fileName, ws.Name, CStr(labels(i)), counts(i + 2)
    Next i
    label = "条件付き書式数"
    ExpectValue fileName, ws.Name, label, ws.Cells.FormatConditions.Count, 100
    ExpectValue fileName, ws.Name, "入力規則セル数", ValidationCount(ws, fileName), 100000
    label = "図形の内訳"
    For Each shp In ws.Shapes
        If shp.Type <> msoComment Then
            shapes = shapes + 1
            If shp.Visible = msoFalse Then hidden = hidden + 1
            If shp.Width < 1 Or shp.Height < 1 Then tiny = tiny + 1
        End If
    Next shp
    ExpectValue fileName, ws.Name, "図形数", shapes, 500, "コメント枠を除くトップレベル。グラフはこの件数にも含む。"
    ExpectValue fileName, ws.Name, "うち非表示", hidden, 0
    ExpectValue fileName, ws.Name, "うち1pt未満", tiny, 0, "Excelが最小サイズに丸めた場合は実数を採用。"
    label = "コメント"
    total = ws.Comments.Count
    threads = ThreadedCount(ws, fileName)
    notes = FAILED
    If IsNumeric(threads) Then
        If CDbl(total) >= CDbl(threads) Then notes = CDbl(total) - CDbl(threads)
    ElseIf CStr(threads) = UNSUPPORTED Then
        notes = UNSUPPORTED
    End If
    ExpectValue fileName, ws.Name, "コメント合計", total, 1000, "現行調査マクロ指定のComments.Count。"
    ExpectValue fileName, ws.Name, "従来のコメント(メモ)", notes, -1, _
        "現行調査マクロ指定のComments.Count－CommentsThreaded.Count。APIの集合定義と実際の作成数の違いに注意。"
    ExpectValue fileName, ws.Name, "スレッドコメント", threads, -1, "CommentsThreaded.Countの実測。返信は除く。"
    ExpectValue fileName, ws.Name, "メモの生コレクション数", total, -1, "ws.Comments.Count。狙ったメモ数の検証用。"
    label = "ピボットテーブル数"
    ExpectValue fileName, ws.Name, label, ws.PivotTables.Count, 10
    label = "グラフ数"
    ExpectValue fileName, ws.Name, label, ws.ChartObjects.Count, 100
    label = "テーブル数"
    ExpectValue fileName, ws.Name, label, ws.ListObjects.Count, 100
    label = "ハイパーリンク数"
    ExpectValue fileName, ws.Name, label, ws.Hyperlinks.Count, 1000
    ExpectValue fileName, ws.Name, "再計算秒", "調査時に計測", 2, CalculationNote(fileName, ws.Name), "実測で判定"
    Exit Sub
Failed:
    AddFailure fileName & " / " & ws.Name, "期待値取得: " & label, Err.Number, Err.Description
    ExpectValue fileName, ws.Name, label, FAILED, -1, "このシートの残りの取得は中断。ほかのシートは継続。"
End Sub

' 保護シートではSpecialCellsを使わず、UsedRangeの右下を読み取る。
Private Function ExpectedLastCell(ByVal ws As Excel.Worksheet, ByRef note As String) As Excel.Range
    Dim used As Excel.Range
    If SheetIsProtected(ws) Then
        Set used = ws.UsedRange
        Set ExpectedLastCell = used.Cells(used.Rows.Count, used.Columns.Count)
        note = "保護を解除せずUsedRangeの右下から取得。保護シートではSpecialCellsを使用しない。"
    Else
        Set ExpectedLastCell = ws.Cells.SpecialCells(xlCellTypeLastCell)
    End If
End Function

' いずれかの保護が有効な場合、SpecialCellsを避ける読取り経路を使用する。
Private Function SheetIsProtected(ByVal ws As Excel.Worksheet) As Boolean
    SheetIsProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
End Function

' 数式の範囲を配列で読み、実数と既知のテスト式ごとの関数フラグを集計する。
Private Function InspectFormulas(ByVal ws As Excel.Worksheet, ByVal fileName As String) As Variant
    Dim result(0 To 9) As Variant, totals(1 To 9) As Double, formulas As Excel.Range
    Dim area As Excel.Range, block As Excel.Range, has As Variant, data As Variant, cell As Excel.Range
    Dim start As Long, rows As Long, batch As Long, columns As Long, r As Long, c As Long, i As Long
    On Error GoTo Failed
    For i = 0 To 9: result(i) = 0: Next i
    has = ws.UsedRange.HasFormula
    If Not IsNull(has) Then
        If has = False Then
            InspectFormulas = result
            Exit Function
        End If
    End If
    If SheetIsProtected(ws) Then
        For Each cell In ws.UsedRange.Cells
            If cell.HasFormula Then
                result(0) = CDbl(result(0)) + 1
                AddFixtureFormula CStr(cell.Formula), totals
            End If
        Next cell
        GoTo Counted
    End If
    Set formulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
    result(0) = formulas.CountLarge
    For Each area In formulas.Areas
        columns = area.Columns.Count
        batch = 8192 \ columns
        If batch < 1 Then batch = 1
        start = 1
        Do While start <= area.Rows.Count
            rows = area.Rows.Count - start + 1
            If rows > batch Then rows = batch
            Set block = area.Cells(start, 1).Resize(rows, columns)
            data = block.Formula
            If block.CountLarge = 1 Then
                AddFixtureFormula CStr(data), totals
            Else
                For r = 1 To rows
                    For c = 1 To columns
                        AddFixtureFormula CStr(data(r, c)), totals
                    Next c
                Next r
            End If
            start = start + rows
        Loop
    Next area
Counted:
    For i = 1 To 9: result(i) = totals(i): Next i
    InspectFormulas = result
    Exit Function
Failed:
    AddFailure fileName & " / " & ws.Name, "数式・揮発性関数の再集計", Err.Number, Err.Description
    For i = 0 To 9: result(i) = FAILED: Next i
    InspectFormulas = result
End Function

' 各生成式と完全一致で照合する独立オラクル。数式文字列内の関数名を検索しない。
Private Function FixtureMask(ByVal formula As String) As Long
    Select Case UCase$(formula)
        Case "=INDIRECT(""D1"")": FixtureMask = 1
        Case "=OFFSET($D$1,0,0)": FixtureMask = 2
        Case "=TODAY()": FixtureMask = 4
        Case "=NOW()": FixtureMask = 8
        Case "=RAND()": FixtureMask = 16
        Case "=RANDBETWEEN(1,100)": FixtureMask = 32
        Case "=CELL(""ADDRESS"",$D$1)": FixtureMask = 64
        Case "=INFO(""SYSTEM"")": FixtureMask = 128
        Case "=TODAY()+NOW()": FixtureMask = 12
        Case "=RAND()+RAND()+RANDBETWEEN(1,100)": FixtureMask = 48
        Case "=SUM(OFFSET($D$1,0,0),INDIRECT(""D1""))": FixtureMask = 3
    End Select
End Function

' 1セルは総数1件とし、同じ関数の繰り返しをビットで重複除外する。
Private Sub AddFixtureFormula(ByVal formula As String, ByRef totals() As Double)
    Dim mask As Long, i As Long
    mask = FixtureMask(formula)
    If mask = 0 Then Exit Sub
    totals(1) = totals(1) + 1
    For i = 0 To 7
        If (mask And CLng(2 ^ i)) <> 0 Then totals(i + 2) = totals(i + 2) + 1
    Next i
End Sub

' 入力規則の件数を再取得する。ほかの1004を0件にしない。
Private Function ValidationCount(ByVal ws As Excel.Worksheet, ByVal fileName As String) As Variant
    Dim cells As Excel.Range, querying As Boolean, number As Long, description As String
    On Error GoTo Failed
    If SheetIsProtected(ws) Then
        ValidationCount = ProtectedValidationCount(ws, fileName)
        Exit Function
    End If
    querying = True
    Set cells = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    querying = False
    ValidationCount = cells.CountLarge
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    If querying And number = 1004 And NoCellsMessage(description) Then
        ValidationCount = 0
    Else
        AddFailure fileName & " / " & ws.Name, "入力規則の再集計", number, description
        ValidationCount = FAILED
    End If
End Function

' 保護された小規模テストシートは、UsedRangeの各セルの入力規則属性を読む。
Private Function ProtectedValidationCount(ByVal ws As Excel.Worksheet, ByVal fileName As String) As Variant
    Dim cell As Excel.Range, kind As Variant, count As Double
    On Error GoTo Failed
    For Each cell In ws.UsedRange.Cells
        kind = ValidationTypeOrNone(cell, fileName, ws.Name)
        If VarType(kind) = vbString Then
            ProtectedValidationCount = FAILED
            Exit Function
        End If
        If CLng(kind) >= 0 Then count = count + 1
    Next cell
    ProtectedValidationCount = count
    Exit Function
Failed:
    AddFailure fileName & " / " & ws.Name, "保護シートの入力規則再集計", Err.Number, Err.Description
    ProtectedValidationCount = FAILED
End Function

' 単一セルのValidation.Typeだけを読み、既知の規則なし1004は-1で返す。
Private Function ValidationTypeOrNone(ByVal cell As Excel.Range, ByVal fileName As String, _
    ByVal sheetName As String) As Variant
    Dim number As Long, description As String
    On Error GoTo Failed
    ValidationTypeOrNone = cell.Validation.Type
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    If number = 1004 And NoValidationMessage(description) Then
        ValidationTypeOrNone = -1
    Else
        AddFailure fileName & " / " & sheetName, "セルの入力規則種類 " & cell.Address(False, False), number, description
        ValidationTypeOrNone = FAILED
    End If
End Function

' Validation.Typeの規則なし時の既知メッセージのみ認め、保護エラー等は正常扱いしない。
Private Function NoValidationMessage(ByVal text As String) As Boolean
    text = LCase$(Trim$(text))
    NoValidationMessage = (text = "application-defined or object-defined error" Or _
        text = "application-defined or object-defined error." Or _
        text = "アプリケーション定義またはオブジェクト定義のエラーです。" Or _
        text = "アプリケーション定義またはオブジェクト定義のエラーです")
End Function

' 日本語・英語の正常な該当なしだけを判定する。
Private Function NoCellsMessage(ByVal text As String) As Boolean
    text = LCase$(Trim$(text))
    NoCellsMessage = (text = "no cells were found." Or text = "no cells were found" Or _
        text = "該当するセルが見つかりません。" Or text = "該当するセルが見つかりません")
End Function

' スレッドAPIがない場合は、通常の未対応状態として返す。
Private Function ThreadedCount(ByVal ws As Object, ByVal fileName As String) As Variant
    Dim number As Long, description As String
    On Error GoTo Failed
    ThreadedCount = ws.CommentsThreaded.Count
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    If number = 438 Then
        ThreadedCount = UNSUPPORTED
    Else
        AddFailure fileName & " / " & ws.Name, "スレッドの再集計", number, description
        ThreadedCount = FAILED
    End If
End Function

' 名前は生成後のRefersTo/Visibleを読む。構造化参照の角括弧だけでは外部扱いしない。
Private Sub InspectNames(ByVal wb As Excel.Workbook, ByVal fileName As String)
    Dim nm As Excel.Name, broken As Long, external As Long, hidden As Long, internalHidden As Long
    Dim nameText As String, refers As String, internal As Boolean, visible As Boolean
    On Error GoTo Failed
    ExpectValue fileName, "ブック", "名前定義数", wb.Names.Count, 1000
    For Each nm In wb.Names
        nameText = nm.Name
        refers = nm.RefersTo
        visible = nm.Visible
        internal = InternalName(nameText)
        If InStr(1, refers, "#REF!", vbTextCompare) > 0 Then broken = broken + 1
        If InStr(1, refers, "[link_target.xlsx]", vbTextCompare) > 0 Then external = external + 1
        If Not visible Then
            If internal Then
                internalHidden = internalHidden + 1
            Else
                hidden = hidden + 1
            End If
        End If
        If Not visible Or InStr(refers, "#REF!") > 0 Or InStr(1, refers, "[link_target.xlsx]", vbTextCompare) > 0 Then
            AddDetail fileName, "ブック", "名前定義", nameText, refers, _
                "表示=" & CStr(visible) & "; 内部名=" & CStr(internal)
        End If
    Next nm
    ExpectValue fileName, "ブック", "うち #REF! の名前定義", broken, 0
    ExpectValue fileName, "ブック", "うち外部参照の名前定義", external, 0
    ExpectValue fileName, "ブック", "うち非表示の名前定義", hidden, 0, "内部名は除外。詳細シートを参照。"
    ExpectValue fileName, "ブック", "内部非表示の名前定義(除外対象)", internalHidden, -1, _
        "heavyではAutoFilterが作成する_FilterDatabaseを含む。実際の名前と表示状態は詳細シートを参照。"
    Exit Sub
Failed:
    AddFailure fileName, "名前定義の再集計", Err.Number, Err.Description
    ExpectValue fileName, "ブック", "名前定義の内訳", FAILED
End Sub

' スコープ付きの内部名を非表示問題から除外する。
Private Function InternalName(ByVal nameText As String) As Boolean
    nameText = LCase$(Mid$(nameText, InStrRev(nameText, "!") + 1))
    InternalName = (nameText Like "_xlfn.*" Or nameText Like "_xlpm.*" Or nameText Like "_xleta.*" Or _
        nameText Like "_xlnm.*" Or nameText = "_filterdatabase")
End Function

' 保存後の外部リンクの実数とリンク先を記録する。
Private Sub InspectLinks(ByVal wb As Excel.Workbook, ByVal fileName As String, _
    ByVal kind As XlLinkType, ByVal label As String)
    Dim links As Variant, i As Long, count As Long
    On Error GoTo Failed
    links = wb.LinkSources(kind)
    If Not IsEmpty(links) Then
        count = UBound(links) - LBound(links) + 1
        For i = LBound(links) To UBound(links)
            AddDetail fileName, "ブック", label, CStr(i), CStr(links(i)), "同じ出力先のlink_target.xlsxを確認。"
        Next i
    End If
    ExpectValue fileName, "ブック", label & "数", count, 0
    Exit Sub
Failed:
    AddFailure fileName, label & "再取得", Err.Number, Err.Description
    ExpectValue fileName, "ブック", label & "数", FAILED
End Sub

' 保存済みの接続・QueryTable・ピボット起動時更新設定を再取得する。
Private Sub InspectConnections(ByVal wb As Excel.Workbook, ByVal fileName As String)
    Dim connection As Excel.WorkbookConnection, ws As Excel.Worksheet, query As Excel.QueryTable
    Dim cache As Excel.PivotCache, i As Long
    On Error GoTo Failed
    ExpectValue fileName, "ブック", "データ接続数", wb.Connections.Count, 0
    For Each connection In wb.Connections
        AddDetail fileName, "ブック", "接続", connection.Name, CStr(connection.Type), "接続先はローカルCSVのみ。"
    Next connection
    For Each ws In wb.Worksheets
        For Each query In ws.QueryTables
            ExpectValue fileName, ws.Name, "QueryTable：ファイルを開くときに更新する", _
                IIf(query.RefreshOnFileOpen, "あり", "なし"), -1, query.Name & " / " & CStr(query.Connection), _
                IIf(query.RefreshOnFileOpen, "要確認", "")
            AddDetail fileName, ws.Name, "QueryTable", query.Name, CStr(query.Connection), _
                "RefreshOnFileOpen=" & CStr(query.RefreshOnFileOpen)
        Next query
    Next ws
    For Each cache In wb.PivotCaches
        i = i + 1
        ExpectValue fileName, "ブック", "ピボットキャッシュ " & CStr(i) & "：ファイルを開くときに更新する", _
            IIf(cache.RefreshOnFileOpen, "あり", "なし"), -1, "保存後のプロパティ値。", IIf(cache.RefreshOnFileOpen, "要確認", "")
    Next cache
    Exit Sub
Failed:
    AddFailure fileName, "接続・起動時更新設定の再取得", Err.Number, Err.Description
    ExpectValue fileName, "ブック", "接続・起動時更新設定", FAILED
End Sub

' 保存したCSVの実際の行数とバイト数を確認する。
Private Sub InspectCSV()
    Dim handle As Integer, opened As Boolean, lineText As String, count As Long, number As Long, description As String
    On Error GoTo Failed
    handle = FreeFile
    Open gFolder & "data.csv" For Input Access Read As #handle
    opened = True
    Do While Not EOF(handle)
        Line Input #handle, lineText
        count = count + 1
    Loop
    Close #handle
    opened = False
    ExpectValue "data.csv", "ファイル", "行数（ヘッダー込み）", count
    ExpectValue "data.csv", "ファイル", "ファイルサイズ(バイト)", FileLen(gFolder & "data.csv")
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    If opened Then CloseFileHandle handle
    AddFailure "data.csv", "保存後のCSV確認", number, description
    ExpectValue "data.csv", "ファイル", "保存後の確認", FAILED
End Sub

' 非負の余分な行列数を返す。
Private Function Nonnegative(ByVal value As Long) As Long
    If value > 0 Then Nonnegative = value
End Function

' 表示状態を文字列に変換する。
Private Function VisibilityText(ByVal value As XlSheetVisibility) As String
    Select Case value
        Case xlSheetVisible: VisibilityText = "表示"
        Case xlSheetHidden: VisibilityText = "非表示"
        Case xlSheetVeryHidden: VisibilityText = "完全非表示"
    End Select
End Function

' 計測せずに、実機の再計算で確認する内容を期待値の備考へ記す。
Private Function CalculationNote(ByVal fileName As String, ByVal sheetName As String) As String
    CalculationNote = "再計算を許可して実測する。別インスタンスではアドイン関数の時間を含まない。"
    If fileName = "test_heavy.xlsx" And sheetName = "重い計算" Then
        CalculationNote = CalculationNote & " このシートの再計算秒が0秒近くにならないこと。" & _
            "SUMPRODUCT式=" & CStr(SUMPRODUCT_FORMULA_COUNT) & "; 各参照行数=" & CStr(SUMPRODUCT_SOURCE_ROWS) & _
            "; 総数式=" & CStr(CALC_FORMULA_COUNT) & "。一般的なPCで数秒を狙うが機種により定数を調整。"
    End If
End Function

' 日付が変わっても負の秒数にならないよう24時間未満の差分を補正する。
Private Function SecondsSince(ByVal started As Double) As Double
    SecondsSince = CDbl(Timer) - started
    If SecondsSince < 0 Then SecondsSince = SecondsSince + 86400#
End Function

' 実測値を保持する。値は実数のまま、目標との差異は備考に追加する。
Private Sub ExpectValue(ByVal fileName As String, ByVal sheetName As String, ByVal item As String, _
    ByVal value As Variant, Optional ByVal threshold As Double = -1, Optional ByVal note As String = "", _
    Optional ByVal overrideJudgment As String = "")
    Dim judgment As String, limit As Variant, difference As String, cause As String
    limit = "-"
    If threshold >= 0 Then
        limit = threshold
        If IsNumeric(value) Then
            If CDbl(value) > threshold Then judgment = "要確認"
        End If
    End If
    If VarType(value) = vbString Then
        If CStr(value) = FAILED Then judgment = "取得失敗"
        If CStr(value) = UNSUPPORTED Then judgment = "非対応"
    End If
    If overrideJudgment <> "" Then judgment = overrideJudgment
    difference = GoalDifference(fileName, sheetName, item, value)
    If difference <> "" Then
        If note <> "" Then note = note & " "
        note = note & difference
        AddFailure fileName & " / " & sheetName, "保存後の目標未達: " & item, 0, difference
    End If
    cause = RelatedFailure(fileName, sheetName, item)
    If cause <> "" Then note = note & " 作成時の理由: " & cause
    gValues.Add Array(fileName, sheetName, item, value, limit, judgment, note)
End Sub

' 該当項目の生成エラーがあれば、APIの実際の理由を期待値の備考にも載せる。
Private Function RelatedFailure(ByVal fileName As String, ByVal sheetName As String, ByVal item As String) As String
    Dim row As Variant, i As Long, feature As String, matches As Boolean
    feature = sheetName
    If sheetName = "ピボット" Then feature = "ピボット・キャッシュ"
    If sheetName = "重い計算" Then feature = "計算式"
    If sheetName = "ブック" Then
        If InStr(item, "スタイル") > 0 Then feature = "スタイル"
        If InStr(item, "名前定義") > 0 Then feature = "名前定義"
        If InStr(item, "ピボットキャッシュ") > 0 Then feature = "ピボット・キャッシュ"
        If InStr(item, "接続") > 0 Then feature = "CSV接続"
    End If
    For i = 1 To gFailures.Count
        row = gFailures(i)
        matches = (CStr(row(0)) = fileName And CStr(row(1)) = feature)
        If sheetName = "コメント" And CStr(row(0)) = fileName Then
            matches = (CStr(row(1)) = "メモ" Or CStr(row(1)) = "スレッドコメント")
        End If
        If sheetName = "ブック" And InStr(item, "内部非表示") > 0 And CStr(row(0)) = fileName Then
            If Left$(CStr(row(1)), Len("内部非表示名")) = "内部非表示名" Then matches = True
        End If
        If matches Then
            If RelatedFailure <> "" Then RelatedFailure = RelatedFailure & "; "
            RelatedFailure = RelatedFailure & CStr(row(1)) & " / " & CStr(row(2)) & " / " & CStr(row(3))
        End If
    Next i
End Function

' 指定の件数と保存後の実数の差を検出し、Excelが保持しなかった項目を明示する。
Private Function GoalDifference(ByVal fileName As String, ByVal sheetName As String, _
    ByVal item As String, ByVal value As Variant) As String
    Dim target As Double, exact As Boolean, key As String
    target = -1
    key = sheetName & "/" & item
    If fileName = "test_heavy.xlsx" Then
        Select Case key
            Case "ブック/うちユーザー定義スタイル数": target = HEAVY_CUSTOM_STYLES
            Case "ブック/名前定義数": target = HEAVY_NORMAL_NAMES
            Case "ブック/うち #REF! の名前定義": target = 3: exact = True
            Case "ブック/うち外部参照の名前定義": target = 2: exact = True
            Case "ブック/うち非表示の名前定義": target = 4: exact = True
            Case "ブック/内部非表示の名前定義(除外対象)": target = 1
            Case "ブック/ピボットキャッシュ数": target = HEAVY_CACHES
            Case "ブック/Excel外部リンク数", "ブック/データ接続数": target = 1
            Case "余分な行/余分な行数": target = 59900
            Case "余分な列/余分な列数": target = 101
            Case "揮発性関数/揮発性関数を含むセル数": target = 67: exact = True
            Case "条件付き書式/条件付き書式数": target = HEAVY_CF
            Case "図形/図形数": target = HEAVY_SHAPES
            Case "図形/うち非表示": target = 5: exact = True
            Case "図形/うち1pt未満": target = 3: exact = True
            Case "コメント/メモの生コレクション数": target = HEAVY_NOTES
            Case "ピボット/ピボットテーブル数": target = HEAVY_PIVOTS
            Case "グラフ/グラフ数": target = HEAVY_CHARTS
            Case "テーブル/テーブル数": target = HEAVY_TABLES
            Case "ハイパーリンク/ハイパーリンク数": target = HEAVY_LINKS
            Case "入力規則/入力規則セル数": target = HEAVY_VALIDATIONS
            Case "重い計算/数式セル数": target = CALC_FORMULA_COUNT: exact = True
        End Select
    ElseIf fileName = "test_edge.xlsx" Then
        exact = True
        Select Case key
            Case "ブック/セルスタイル数": target = EDGE_STYLES
            Case "ブック/名前定義数": target = EDGE_NAMES
            Case "ブック/ピボットキャッシュ数": target = EDGE_CACHES
            Case "ブック/うち #REF! の名前定義", "ブック/うち外部参照の名前定義", _
                "ブック/うち非表示の名前定義", "ブック/Excel外部リンク数": target = 0
            Case "余分な行/余分な行数": target = 1000
            Case "余分な列/余分な列数": target = 100
            Case "条件付き書式/条件付き書式数": target = EDGE_CF
            Case "図形/図形数": target = EDGE_SHAPES
            Case "コメント/メモの生コレクション数": target = EDGE_NOTES
            Case "ピボット/ピボットテーブル数": target = EDGE_PIVOTS
            Case "グラフ/グラフ数": target = EDGE_CHARTS
            Case "テーブル/テーブル数": target = EDGE_TABLES
            Case "ハイパーリンク/ハイパーリンク数": target = EDGE_LINKS
            Case "入力規則/入力規則セル数": target = EDGE_VALIDATIONS
            Case "重い計算/数式セル数": target = EDGE_FORMULAS
            Case "誤検出/揮発性関数を含むセル数": target = 0
        End Select
    End If
    If target < 0 Then Exit Function
    If Not IsNumeric(value) Then
        GoalDifference = "保存後の件数を確定できません。作成記録を参照。狙い値=" & CStr(target)
    ElseIf CDbl(value) < target Or (exact And CDbl(value) <> target) Then
        GoalDifference = "狙い値=" & CStr(target) & "、保存後の実数=" & CStr(value) & _
            "。Excelの保持・自動追加・最小サイズ・API対応等で一致しませんでした。作成記録と詳細を参照。"
    End If
End Function

' 保存後に実際に保持された参照・名前・接続を別シートへ残す。
Private Sub AddDetail(ByVal fileName As String, ByVal sheetName As String, ByVal kind As String, _
    ByVal nameText As String, ByVal value As String, ByVal note As String)
    gDetails.Add Array(fileName, sheetName, kind, nameText, value, note)
End Sub

' エラーの値を受け取り、後で一覧とメッセージに出力する。
Private Sub AddFailure(ByVal place As String, ByVal item As String, ByVal number As Long, ByVal description As String)
    gFailures.Add Array(place, item, number, description)
End Sub

' 再計算と各ファイルの確認観点をexpected.xlsxにも収録する。
Private Sub BuildChecks()
    gChecks.Add Array("test_clean.xlsx", "全体", "再計算いいえで各項目の要確認がないこと", "サイズと時間は環境依存の実測。")
    gChecks.Add Array("test_heavy.xlsx", "各専用シート", "各しきい値を1枚単独で超過すること", "期待値シートの実数・判定と比較。失敗項目は作成記録を参照。")
    gChecks.Add Array("test_heavy.xlsx", "余分な行 / 余分な列", "Z60000とDA100を別々に検出すること", "A1:D100が実データ。余分な行59900・列101を別シートで確認。")
    gChecks.Add Array("test_heavy.xlsx", "揮発性関数", "総セル数67。1セル内の同じ関数は1件", "関数別は2,4,6,8,10,12,13,15。実際に保存された式を再照合した値と比較。")
    gChecks.Add Array("test_heavy.xlsx", "コメント", "CommentsとCommentsThreadedの実数を比較", "メモの生コレクションと、現行調査マクロ指定の差分を両方記載。")
    gChecks.Add Array("test_heavy.xlsx", "重い計算", "再計算を許可したとき、再計算秒が0秒近くにならないこと", _
        "一般的なPCで数秒を狙う。CALC_FORMULA_COUNT / SUMPRODUCT_FORMULA_COUNT / SUMPRODUCT_SOURCE_ROWSで調整。生成中は計算しない。")
    gChecks.Add Array("test_heavy.xlsx", "重い計算", "揮発性関数セル数0、総数式100000超", "総数の定数を変更した場合は実数に従う。")
    gChecks.Add Array("test_heavy.xlsx", "外部リンク / CSV接続", "同じフォルダの補助ファイルだけを参照し、CSVの起動時更新があり", "生成用の明示Refreshは、このローカルCSVの最初の取込だけ。")
    gChecks.Add Array("test_edge.xlsx", "境界用各シート", "値がしきい値ちょうどなら要確認にならないこと", "スタイル総数2000、名前1000、ルール100、図形500、コメント1000等。")
    gChecks.Add Array("test_edge.xlsx", "誤検出", "揮発性関数セル数0、名前の外部参照0", "文字列、OFFSET(シート、NOW()列、RANDARRAY、MYOFFSET、構造化参照。")
    gChecks.Add Array("test_edge.xlsx", "空 / 非表示 / 完全非表示 / 保護 / グラフシート", "空シート0件、表示状態の一致、保護シートが読み取り可能", "グラフシートは全シート数に含み、調査のワークシート別には出ない。")
    gChecks.Add Array("test_legacy.xls", "形式", "旧形式の注意喚起が要確認", "Excel 97-2003形式。マクロを含まない。")
    gChecks.Add Array("全ファイル", "期待値", "作成試行回数ではなく、保存後に開き直した実数で比較", "再計算・サイズ・開く時間以外のしきい値は現行HeavyBookCheckの既定値。")
End Sub

' バッファを4シートへ書き出し、最後の1ファイルとして保存する。
Private Function SaveExpected() As Boolean
    Dim ws As Excel.Worksheet, i As Long
    On Error GoTo Failed
    AssertNewPath gFolder & "expected.xlsx"
    gExpected.Worksheets(1).Name = "期待値"
    Set ws = gExpected.Worksheets(1)
    WriteCollection ws, Array("ファイル", "シート・範囲", "調査項目", "保存後の実測期待値", "しきい値(超過)", "期待判定", "備考"), gValues, "A:C,F:G"
    Set ws = AddSheet(gExpected, "名前リンク接続")
    WriteCollection ws, Array("ファイル", "シート・範囲", "種類", "名前", "実際の参照先・値", "備考"), gDetails, "A:F"
    Set ws = AddSheet(gExpected, "確認観点")
    WriteCollection ws, Array("ファイル", "対象", "確認すること", "補足"), gChecks, "A:D"
    Set ws = AddSheet(gExpected, "作成記録")
    WriteCollection ws, Array("場所", "失敗・未対応項目", "エラー番号", "内容・理由"), gFailures, "A:B,D:D"
    For Each ws In gExpected.Worksheets
        With ws.UsedRange
            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(221, 235, 247)
            .Columns.ColumnWidth = 22
            If .Rows.Count > 1 Then .AutoFilter
        End With
    Next ws
    gExpected.Worksheets("期待値").Columns("G").ColumnWidth = 70
    SaveNewBook gExpected, "expected.xlsx", xlOpenXMLWorkbook
    SaveExpected = True
    Exit Function
Failed:
    AddFailure "expected.xlsx", "期待値ブックの保存", Err.Number, Err.Description
End Function

' 行コレクションを一括出力する。参照式を含む文字列列は評価しない。
Private Sub WriteCollection(ByVal ws As Excel.Worksheet, ByVal headers As Variant, _
    ByVal records As Collection, ByVal textColumns As String)
    Dim data() As Variant, i As Long, c As Long, row As Variant
    Dim columns As Long, value As Variant
    columns = UBound(headers) + 1
    ReDim data(1 To records.Count + 1, 1 To columns)
    For c = 1 To columns
        data(1, c) = headers(c - 1)
    Next c
    For i = 1 To records.Count
        row = records(i)
        For c = 1 To columns
            value = row(c - 1)
            If VarType(value) = vbString Then value = Left$(CStr(value), 32767)
            data(i + 1, c) = value
        Next c
    Next i
    ws.Range(textColumns).NumberFormat = "@"
    ws.Range("A1").Resize(records.Count + 1, columns).Value2 = data
End Sub

' 結果ブック以外をすべて閉じ、重い式が設定復元で計算されることを防ぐ。
Private Function CloseOtherBooks() As Boolean
    Dim i As Long, wb As Excel.Workbook, closed As Boolean
    On Error GoTo Failed
    closed = True
    If gExcel Is Nothing Then
        CloseOtherBooks = True
        Exit Function
    End If
    For i = gExcel.Workbooks.Count To 1 Step -1
        Set wb = gExcel.Workbooks(i)
        If Not wb Is gExpected Then
            If Not CloseOneBook(wb) Then closed = False
        End If
    Next i
    CloseOtherBooks = closed
    Exit Function
Failed:
    AddFailure "後始末", "生成・再取得ブックを閉じる", Err.Number, Err.Description
End Function

' 対象を再保存せず閉じる。ほかの項目の処理を止めない。
Private Function CloseOneBook(ByVal wb As Excel.Workbook) As Boolean
    On Error GoTo Failed
    wb.Close SaveChanges:=False
    CloseOneBook = True
    Exit Function
Failed:
    AddFailure "後始末", "ブックを閉じる", Err.Number, Err.Description
End Function

' 専用Excelを終了する。終了できない異常時は非表示で残さず表示する。
Private Sub CloseAndQuit()
    On Error GoTo Failed
    If gExcel Is Nothing Then Exit Sub
    gExcel.DisplayAlerts = False
    gExcel.Quit
    Exit Sub
Failed:
    AddFailure "後始末", "専用Excelの終了", Err.Number, Err.Description
    ShowOwnedExcel
End Sub

' 専用インスタンスの終了に失敗した場合の最終表示を試みる。
Private Sub ShowOwnedExcel()
    On Error GoTo Failed
    gExcel.AutomationSecurity = msoAutomationSecurityByUI
    gExcel.Visible = True
    gExcel.UserControl = True
    Exit Sub
Failed:
    AddFailure "後始末", "専用Excelの表示", Err.Number, Err.Description
End Sub

' 失敗一覧は長くても切り捨てず、メッセージを複数ページに分けて表示する。
Private Sub ShowFailureList(ByVal prefix As String, ByVal reportSaved As Boolean)
    Dim text As String, row As Variant, i As Long, start As Long
    text = prefix & vbCrLf
    If reportSaved Then
        text = text & "expected.xlsxの作成記録も確認してください。終了処理の失敗はこの一覧を確認してください。" & vbCrLf
    Else
        text = text & "期待値ブックは保存されていません。この一覧を確認してください。" & vbCrLf
    End If
    For i = 1 To gFailures.Count
        row = gFailures(i)
        text = text & vbCrLf & CStr(i) & ". " & CStr(row(0)) & " / " & CStr(row(1)) & _
            " / " & CStr(row(2)) & " / " & CStr(row(3))
    Next i
    start = 1
    Do While start <= Len(text)
        MsgBox Mid$(text, start, 900), vbExclamation, "テスト生成：失敗・未対応一覧"
        start = start + 900
    Loop
End Sub
