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
' --- SearchShapeForm のコード ---

' 検索キーワードが変更されたかを判断するための変数
Private lastSearchTerm As String
' 見つかった図形を格納するコレクション
Private foundShapes As Collection
' 次に表示する図形のインデックス
Private currentShapeIndex As Long

' 「終了」ボタンが押されたときの処理
Private Sub btnClose_Click()
    Unload Me
End Sub

' テキストボックスでキーが押されたときの処理（エラー対策版）
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
        Dim hasTextFrameProperty As Boolean
        
        ' アクティブなシートのすべての図形をループ
        For Each shp In ActiveSheet.Shapes
            
            '--- ▼ここからが修正部分▼ ---
            hasTextFrameProperty = False
            On Error Resume Next ' 一時的にエラーを無視する
            hasTextFrameProperty = shp.HasTextFrame ' このプロパティが存在するか試す
            On Error GoTo 0      ' エラー無視をすぐに解除する
            '--- ▲ここまでが修正部分▲ ---

            ' HasTextFrameプロパティがあり、かつテキストが含まれている場合のみ処理
            If hasTextFrameProperty Then
                If shp.TextFrame2.HasText Then
                    ' テキスト内に検索キーワードが含まれているかチェック（大文字・小文字を区別しない）
                    If InStr(1, shp.TextFrame.Characters.Text, searchTerm, vbTextCompare) > 0 Then
                        foundShapes.Add shp
                    End If
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


' フォームが閉じられるときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 変数を解放
    Set foundShapes = Nothing
End Sub
