VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchShapeForm 
   Caption         =   "図形内テキスト検索"
   ClientHeight    =   4110
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
' 検索キーワードが変更されたかを判断するための変数
Private lastSearchTerm As String
' 見つかった図形を格納するコレクション
Private foundShapes As Collection
' 次に表示する図形のインデックス
Private currentShapeIndex As Long


' 図形からテキストを安全に取得するための専用関数
Private Function GetShapeText(ByVal targetShape As Shape) As String
    On Error Resume Next ' この関数内で発生したエラーはすべて無視する
    GetShapeText = "" ' 初期値を設定
    
    ' テキストフレームがあり、かつ、テキストが含まれている場合にその内容を返す
    If targetShape.HasTextFrame Then
        If targetShape.TextFrame2.HasText Then
            GetShapeText = targetShape.TextFrame.Characters.Text
        End If
    End If
    
    On Error GoTo 0 ' エラー処理を元に戻す
End Function


' 図形を再帰的に検索するためのプロシージャ
Private Sub SearchShapesRecursive(ByVal shapesToSearch As Object, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next ' GroupItemsなどでエラーが出ることがあるため

    Dim shp As Shape
    Dim shapeText As String

    For Each shp In shapesToSearch
        ' 図形がグループの場合、その中のアイテムに対して再度このプロシージャを呼び出す（再帰処理）
        If shp.Type = msoGroup Then
            SearchShapesRecursive shp.GroupItems, searchTerm, results
        Else
            ' グループでない場合、通常通りテキストを検索
            shapeText = GetShapeText(shp)
            If Len(shapeText) > 0 Then
                If InStr(1, shapeText, searchTerm, vbTextCompare) > 0 Then
                    results.Add shp
                End If
            End If
        End If
    Next shp
    
    On Error GoTo 0
End Sub


' テキストボックスでキーが押されたときの処理（最終版）
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    ' Enterキーが押された場合のみ実行
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    ' Enterキーの入力を無効化（「ポン」という音を防ぐ）
    KeyCode = 0
    
    '--- ▼ここからが修正部分▼ ---
    ' btnSearch_Clickプロシージャを呼び出す
    Call btnSearch_Click
    '--- ▲ここまでが修正部分▲ ---
    
End Sub

'--- ▼ここからが新規追加部分▼ ---
' 「次を検索」ボタンが押されたときの処理（元のKeyDownイベントから処理を移動）
Private Sub btnSearch_Click()

    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' 検索キーワードが空なら何もしない
    If Len(searchTerm) = 0 Then Exit Sub
    
    ' --- 検索処理 ---
    ' 検索キーワードが変わったか、初めての検索の場合
    If lastSearchTerm <> searchTerm Or foundShapes Is Nothing Then
        lastSearchTerm = searchTerm
        Set foundShapes = New Collection
        currentShapeIndex = 0
        
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
    End If
    
    ' --- 検索結果の表示 ---
    If foundShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        
        ' 図形がどのシート上にあるかを直接調べて、そのシートをアクティブにする
        On Error Resume Next ' 特殊なオブジェクトでTopLeftCellが取得できない場合に備える
        targetShape.TopLeftCell.Worksheet.Activate
        On Error GoTo 0
        
        ' --- スクロール処理 ---
        On Error Resume Next
        Application.Goto Reference:=targetShape, Scroll:=True
        If Err.Number <> 0 Then
            Err.Clear
            targetShape.Select
            With ActiveWindow
                .ScrollRow = targetShape.TopLeftCell.Row
                .ScrollColumn = targetShape.TopLeftCell.Column
            End With
        End If
        On Error GoTo 0
        
    Else
        Beep
    End If

End Sub


' 「置換」ボタンが押されたときの処理
Private Sub btnReplace_Click()
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    ' 検索キーワードが空、または検索結果がない場合は何もしない
    If Len(searchTerm) = 0 Or foundShapes Is Nothing Or foundShapes.Count = 0 Then
        Beep
        Exit Sub
    End If

    ' 現在選択されている図形を取得
    Dim targetShape As Shape
    Set targetShape = foundShapes(currentShapeIndex)

    ' テキストを置換
    Dim originalText As String, newText As String
    originalText = GetShapeText(targetShape)
    ' InStrで大文字/小文字を区別せずに検索し、見つかった部分をReplaceで置換
    If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
        newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
        targetShape.TextFrame.Characters.Text = newText
    End If
    
    ' 次の図形を検索して表示
    Call btnSearch_Click

End Sub


' 「すべて置換」ボタンが押されたときの処理
Private Sub btnReplaceAll_Click()
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    ' 検索キーワードが空の場合は何もしない
    If Len(searchTerm) = 0 Then Exit Sub

    ' --- 検索がまだ実行されていない場合は、まず検索を実行 ---
    If foundShapes Is Nothing Or lastSearchTerm <> searchTerm Then
        lastSearchTerm = searchTerm
        Set foundShapes = New Collection
        currentShapeIndex = 0
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
    End If

    ' --- 置換処理 ---
    If foundShapes.Count > 0 Then
        Dim shp As Shape
        Dim replacedCount As Long
        replacedCount = 0

        On Error Resume Next
        For Each shp In foundShapes
            Dim originalText As String, newText As String
            originalText = GetShapeText(shp)
            
            ' 置換対象のテキストが含まれている場合のみ処理
            If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
                newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
                shp.TextFrame.Characters.Text = newText
                replacedCount = replacedCount + 1
            End If
        Next shp
        On Error GoTo 0

        MsgBox replacedCount & "個の項目を置換しました。", vbInformation
        
        ' 置換が終わったので検索結果をクリア
        Set foundShapes = Nothing
        lastSearchTerm = ""
        currentShapeIndex = 0

    Else
        MsgBox "置換対象の図形が見つかりませんでした。", vbExclamation
    End If

End Sub
'--- ▲ここまでが新規追加部分▲ ---


' 「終了」ボタンが押されたときの処理
Private Sub btnClose_Click()
    Unload Me
End Sub

' フォームが閉じられるときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundShapes = Nothing
End Sub
