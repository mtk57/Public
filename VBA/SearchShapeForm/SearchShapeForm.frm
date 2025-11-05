VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchShapeForm 
   Caption         =   "テキスト検索"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   OleObjectBlob   =   "SearchShapeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SearchShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VER = "1.2.0" ' バージョンを更新

' --- 変数定義 ---
Private lastSearchTerm As String
Private foundItems As Collection
Private currentShapeIndex As Long
Private regex As Object

' フォームが初期化されたときの処理
Private Sub UserForm_Initialize()
    ' フォームのキャプションにバージョンを追加
    Me.Caption = "図形/セル内の文字列の検索/置換 (ver " & VER & ")"
    
    ' ラジオボタンのデフォルトを「現在のシートのみ」に設定
    Me.optCurrentSheet.Value = True
    
    ' チェックボックスのデフォルト値を設定
    Me.chkSearchInCells.Value = True ' デフォルトでセルも検索対象にする
    Me.chkUseRegex.Value = False     ' デフォルトで正規表現は使用しない
    
    ' ★追加: 新しいチェックボックスのデフォルト値を設定
    Me.chkPartialMatch.Value = True  ' デフォルトで部分一致をオンにする
    Me.ChkCase.Value = False          ' デフォルトで大文字小文字を区別しない
    
    ' 正規表現オブジェクトを初期化
    Set regex = CreateObject("VBScript.RegExp")
End Sub

' --- ヘルパー関数 ---

' 図形からテキストを安全に取得するための専用関数（変更なし）
Private Function GetShapeText(ByVal targetShape As Shape) As String
    On Error Resume Next
    GetShapeText = ""
    If targetShape.HasTextFrame Then
        If targetShape.TextFrame2.HasText Then
            GetShapeText = targetShape.TextFrame.Characters.Text
        End If
    End If
    On Error GoTo 0
End Function

' ★更新: テキストが検索条件に一致するかを判定する関数
Private Function IsMatch(ByVal inputText As String, ByVal searchTerm As String) As Boolean
    If Len(inputText) = 0 Or Len(searchTerm) = 0 Then
        IsMatch = False
        Exit Function
    End If

    Dim effectiveSearchTerm As String
    
    If Me.chkUseRegex.Value Then
        ' --- 正規表現を使用する場合 ---
        With regex
            ' 大文字小文字の区別を設定
            .IgnoreCase = Not Me.ChkCase.Value
            
            ' 部分一致でない場合、完全一致するようにパターンを変更
            If Not Me.chkPartialMatch.Value Then
                effectiveSearchTerm = "^" & searchTerm & "$"
            Else
                effectiveSearchTerm = searchTerm
            End If
            
            .Pattern = effectiveSearchTerm
            .Global = True
            .MultiLine = True
            IsMatch = .Test(inputText)
        End With
    Else
        ' --- 通常のテキスト検索の場合 ---
        Dim compareMethod As VbCompareMethod
        ' 大文字小文字の区別を設定
        If Me.ChkCase.Value Then
            compareMethod = vbBinaryCompare ' 区別する
        Else
            compareMethod = vbTextCompare   ' 区別しない
        End If

        ' 部分一致かどうかで処理を分岐
        If Me.chkPartialMatch.Value Then
            ' 部分一致検索
            IsMatch = (InStr(1, inputText, searchTerm, compareMethod) > 0)
        Else
            ' 完全一致検索
            IsMatch = (StrComp(inputText, searchTerm, compareMethod) = 0)
        End If
    End If
End Function


' ★更新: 正規表現や大文字/小文字区別を考慮したテキスト置換関数
Private Function DoReplace(ByVal inputText As String, ByVal searchTerm As String, ByVal replaceTerm As String) As String
    Dim effectiveSearchTerm As String
    
    If Me.chkUseRegex.Value Then
        ' --- 正規表現を使用する場合 ---
        With regex
            ' 大文字小文字の区別を設定
            .IgnoreCase = Not Me.ChkCase.Value
            
            ' 部分一致でない場合、検索パターンを完全一致に変更
            If Not Me.chkPartialMatch.Value Then
                 effectiveSearchTerm = "^" & searchTerm & "$"
            Else
                 effectiveSearchTerm = searchTerm
            End If

            .Pattern = effectiveSearchTerm
            .Global = True
            .MultiLine = True
            DoReplace = .Replace(inputText, replaceTerm)
        End With
    Else
        ' --- 通常のテキスト検索の場合 ---
        Dim compareMethod As VbCompareMethod
        If Me.ChkCase.Value Then
            compareMethod = vbBinaryCompare ' 区別する
        Else
            compareMethod = vbTextCompare   ' 区別しない
        End If

        If Me.chkPartialMatch.Value Then
            ' 部分一致の置換
            DoReplace = Replace(inputText, searchTerm, replaceTerm, 1, -1, compareMethod)
        Else
            ' 完全一致の場合、テキスト全体が検索語と一致する場合のみ置換
            If StrComp(inputText, searchTerm, compareMethod) = 0 Then
                DoReplace = replaceTerm
            Else
                DoReplace = inputText ' 一致しない場合は元のテキストを返す
            End If
        End If
    End If
End Function

' --- 検索プロシージャ ---

' 図形を再帰的に検索するためのプロシージャ（変更なし）
Private Sub SearchShapesRecursive(ByVal shapesToSearch As Object, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next
    Dim shp As Shape
    Dim shapeText As String
    For Each shp In shapesToSearch
        If shp.Type = msoGroup Then
            SearchShapesRecursive shp.GroupItems, searchTerm, results
        Else
            shapeText = GetShapeText(shp)
            If IsMatch(shapeText, searchTerm) Then
                results.Add shp
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

' セルを検索するためのプロシージャ（変更なし）
Private Sub SearchCells(ByVal sheetToSearch As Worksheet, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next
    Dim cell As Range
    For Each cell In sheetToSearch.UsedRange.Cells
        If Not cell.HasFormula Then
            If IsMatch(cell.Text, searchTerm) Then
                results.Add cell
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

' 検索を実行する共通プロシージャ（変更なし）
Private Sub ExecuteSearch()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' 検索結果を初期化
    Set foundItems = New Collection
    currentShapeIndex = 0
    lastSearchTerm = searchTerm

    Dim ws As Worksheet
    If Me.optAllSheets.Value = True Then
        ' すべてのシートを検索
        For Each ws In ActiveWorkbook.Worksheets
            SearchShapesRecursive ws.Shapes, searchTerm, foundItems
            If Me.chkSearchInCells.Value Then
                SearchCells ws, searchTerm, foundItems
            End If
        Next ws
    Else
        ' 現在のシートのみ検索
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundItems
        If Me.chkSearchInCells.Value Then
            SearchCells ActiveSheet, searchTerm, foundItems
        End If
    End If
End Sub

' --- イベントハンドラ ---

' テキストボックスでEnterキーが押されたときの処理（変更なし）
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    Call btnSearch_Click
End Sub

' 「次を検索」ボタンが押されたときの処理（変更なし）
Private Sub btnSearch_Click()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    If Len(searchTerm) = 0 Then Exit Sub
    
    ' --- 検索処理 ---
    If lastSearchTerm <> searchTerm Or foundItems Is Nothing Then
        Call ExecuteSearch
    End If
    
    ' --- 検索結果の表示 ---
    If foundItems.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundItems.Count Then currentShapeIndex = 1
        
        Dim targetItem As Object
        Set targetItem = foundItems(currentShapeIndex)
        
        Dim targetSheet As Worksheet
        On Error Resume Next
        If TypeName(targetItem) = "Range" Then
            Set targetSheet = targetItem.Worksheet
        Else ' Shape
            Set targetSheet = targetItem.TopLeftCell.Worksheet
        End If
        targetSheet.Activate
        
        ' --- スクロール処理 ---
        Application.GoTo Reference:=targetItem, Scroll:=True
        If Err.Number <> 0 Then
            Err.Clear
            targetItem.Select
            With ActiveWindow
                If TypeName(targetItem) = "Range" Then
                    .ScrollRow = targetItem.Row
                    .ScrollColumn = targetItem.Column
                Else ' Shape
                    .ScrollRow = targetItem.TopLeftCell.Row
                    .ScrollColumn = targetItem.TopLeftCell.Column
                End If
            End With
        End If
        On Error GoTo 0
    Else
        Beep
    End If

    AppActivate Me.Caption
    Me.txtSearch.SetFocus
End Sub

' 「置換」ボタンが押されたときの処理（変更なし）
Private Sub btnReplace_Click()
    If MsgBox("現在の項目を置換しますか？", vbYesNo + vbQuestion, "置換の確認") = vbNo Then Exit Sub
    
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Or foundItems Is Nothing Or foundItems.Count = 0 Then
        Beep
        Exit Sub
    End If

    Dim targetItem As Object
    Set targetItem = foundItems(currentShapeIndex)

    Dim originalText As String, newText As String
    If TypeName(targetItem) = "Range" Then
        originalText = CStr(targetItem.Value)
        newText = DoReplace(originalText, searchTerm, replaceTerm)
        targetItem.Value = newText
    Else ' Shape
        originalText = GetShapeText(targetItem)
        If IsMatch(originalText, searchTerm) Then
            newText = DoReplace(originalText, searchTerm, replaceTerm)
            targetItem.TextFrame.Characters.Text = newText
        End If
    End If
    
    Call btnSearch_Click
End Sub

' 「すべて置換」ボタンが押されたときの処理（変更なし）
Private Sub btnReplaceAll_Click()
    Dim scopeText As String
    If Me.optAllSheets.Value = True Then scopeText = "すべてのシート" Else scopeText = "現在のシート"

    Dim confirmMsg As String
    confirmMsg = "検索範囲：「" & scopeText & "」" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "「" & Me.txtSearch.Text & "」をすべて「" & Me.txtReplace.Text & "」に置換します。" & vbCrLf
    confirmMsg = confirmMsg & "この操作は元に戻せません。よろしいですか？"
    If MsgBox(confirmMsg, vbYesNo + vbExclamation, "すべての置換の確認") = vbNo Then Exit Sub

    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text
    If Len(searchTerm) = 0 Then Exit Sub

    If foundItems Is Nothing Or lastSearchTerm <> searchTerm Then Call ExecuteSearch

    If foundItems.Count > 0 Then
        Dim item As Object
        Dim replacedCount As Long
        replacedCount = 0
        On Error Resume Next
        For Each item In foundItems
            Dim originalText As String, newText As String
            If TypeName(item) = "Range" Then
                originalText = CStr(item.Value)
                If IsMatch(originalText, searchTerm) Then
                    item.Value = DoReplace(originalText, searchTerm, replaceTerm)
                    replacedCount = replacedCount + 1
                End If
            Else ' Shape
                originalText = GetShapeText(item)
                If IsMatch(originalText, searchTerm) Then
                    item.TextFrame.Characters.Text = DoReplace(originalText, searchTerm, replaceTerm)
                    replacedCount = replacedCount + 1
                End If
            End If
        Next item
        On Error GoTo 0

        MsgBox replacedCount & "個の項目を置換しました。", vbInformation
        
        Set foundItems = Nothing
        lastSearchTerm = ""
        currentShapeIndex = 0
    Else
        MsgBox "置換対象の項目が見つかりませんでした。", vbExclamation
    End If
End Sub

' ★追加: 新しいチェックボックスのクリックイベントで検索結果をリセット
Private Sub chkCase_Click()
    Set foundItems = Nothing
End Sub

Private Sub chkPartialMatch_Click()
    Set foundItems = Nothing
End Sub

' オプションが変更されたら、検索結果をリセットする
Private Sub chkSearchInCells_Click()
    Set foundItems = Nothing
End Sub

Private Sub chkUseRegex_Click()
    Set foundItems = Nothing
End Sub

Private Sub optAllSheets_Click()
    Set foundItems = Nothing
End Sub

Private Sub optCurrentSheet_Click()
    Set foundItems = Nothing
End Sub

' 「終了」ボタンが押されたときの処理（変更なし）
Private Sub btnClose_Click()
    Unload Me
End Sub

' フォームが閉じられるときの処理（変更なし）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundItems = Nothing
    Set regex = Nothing
End Sub
