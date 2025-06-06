VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchShapeForm 
   Caption         =   "図形内テキスト検索"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SearchShapeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SearchShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- SearchShapeForm のコード（グループ化対応版） ---

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


'--- ▼ここからが新規追加部分▼ ---
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
'--- ▲ここまでが新規追加部分▲ ---


' テキストボックスでキーが押されたときの処理
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    ' Enterキーが押された場合のみ実行
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    ' Enterキーの入力を無効化（「ポン」という音を防ぐ）
    KeyCode = 0
    
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
        
        '--- ▼検索ロジックを再帰呼び出しに変更▼ ---
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
        '--- ▲検索ロジックを再帰呼び出しに変更▲ ---
    End If
    
    ' --- 検索結果の表示 ---
    If foundShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        
        targetShape.Parent.Activate
        
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


' 「終了」ボタンが押されたときの処理
Private Sub btnClose_Click()
    Unload Me
End Sub

' フォームが閉じられるときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundShapes = Nothing
End Sub
