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
' --- SearchShapeForm のコード（改善版） ---

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
        
        Dim shp As Shape
        Dim shapeText As String
        
        ' アクティブなシートのすべての図形をループ
        For Each shp In ActiveSheet.Shapes
            ' 専用関数を使って、安全に図形のテキストを取得
            shapeText = GetShapeText(shp)
            
            ' 取得したテキスト内に検索キーワードが含まれているかチェック
            If Len(shapeText) > 0 Then
                If InStr(1, shapeText, searchTerm, vbTextCompare) > 0 Then
                    foundShapes.Add shp
                End If
            End If
        Next shp
    End If
    
    ' --- 検索結果の表示 ---
    If foundShapes.Count > 0 Then
        ' 次の図形へインデックスを進める
        currentShapeIndex = currentShapeIndex + 1
        ' インデックスが図形の数を超えたら1に戻る（ループする）
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        ' 対象の図形を選択して、そこまでスクロール
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        targetShape.Parent.Activate
        Application.Goto Reference:=targetShape, Scroll:=True
        
    Else
        ' 見つからなかった場合
        Beep
    End If
    
End Sub


' 「終了」ボタンが押されたときの処理
Private Sub btnClose_Click()
    Unload Me
End Sub

' フォームが閉じられるときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 変数を解放
    Set foundShapes = Nothing
End Sub
