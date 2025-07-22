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
Const VER = "1.0.1"

' 検索キーワードが変更されたかを判断するための変数
Private lastSearchTerm As String
' 見つかった図形を格納するコレクション
Private foundShapes As Collection
' 次に表示する図形のインデックス
Private currentShapeIndex As Long

' フォームが初期化されたときの処理
Private Sub UserForm_Initialize()
    ' フォームのキャプションにバージョンを追加
    Me.Caption = "図形内の文字列の検索/置換 (ver " & VER & ")"
    
    ' ラジオボタンのデフォルトを「現在のシートのみ」に設定
    Me.optCurrentSheet.Value = True
End Sub


' 図形からテキストを安全に取得するための専用関数
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


' 図形を再帰的に検索するためのプロシージャ
Private Sub SearchShapesRecursive(ByVal shapesToSearch As Object, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next

    Dim shp As Shape
    Dim shapeText As String

    For Each shp In shapesToSearch
        If shp.Type = msoGroup Then
            SearchShapesRecursive shp.GroupItems, searchTerm, results
        Else
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


' 検索を実行する共通プロシージャ
Private Sub ExecuteSearch()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' 検索結果を初期化
    Set foundShapes = New Collection
    currentShapeIndex = 0
    lastSearchTerm = searchTerm

    If Me.optAllSheets.Value = True Then
        ' すべてのシートを検索
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            SearchShapesRecursive ws.Shapes, searchTerm, foundShapes
        Next ws
    Else
        ' 現在のシートのみ検索
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
    End If
End Sub


' テキストボックスでキーが押されたときの処理
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    Call btnSearch_Click
End Sub


' 「次を検索」ボタンが押されたときの処理
Private Sub btnSearch_Click()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    If Len(searchTerm) = 0 Then Exit Sub
    
    ' --- 検索処理 ---
    ' 検索キーワードが変わったか、初めての検索(foundShapesが空)の場合に共通検索処理を呼び出す
    If lastSearchTerm <> searchTerm Or foundShapes Is Nothing Then
        Call ExecuteSearch
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
        On Error Resume Next
        targetShape.TopLeftCell.Worksheet.Activate
        On Error GoTo 0
        
        ' --- スクロール処理 ---
        On Error Resume Next
        Application.GoTo Reference:=targetShape, Scroll:=True
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

    ' ★変更点: AppActivateでフォームを強制的にアクティブにし、テキストボックスにフォーカスを戻す
    On Error Resume Next ' AppActivateが稀に失敗する場合があるためエラーを無視する
    AppActivate Me.Caption
    On Error GoTo 0
    Me.txtSearch.SetFocus
End Sub


' 「置換」ボタンが押されたときの処理
Private Sub btnReplace_Click()
    ' 確認ダイアログの表示
    If MsgBox("現在の図形のテキストを置換しますか？", vbYesNo + vbQuestion, "置換の確認") = vbNo Then
        Exit Sub
    End If
    
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Or foundShapes Is Nothing Or foundShapes.Count = 0 Then
        Beep
        Exit Sub
    End If

    Dim targetShape As Shape
    Set targetShape = foundShapes(currentShapeIndex)

    Dim originalText As String, newText As String
    originalText = GetShapeText(targetShape)
    If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
        newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
        targetShape.TextFrame.Characters.Text = newText
    End If
    
    Call btnSearch_Click
End Sub


' 「すべて置換」ボタンが押されたときの処理
Private Sub btnReplaceAll_Click()
    Dim scopeText As String
    If Me.optAllSheets.Value = True Then
        scopeText = "すべてのシート"
    Else
        scopeText = "現在のシート"
    End If

    Dim confirmMsg As String
    confirmMsg = "検索範囲：「" & scopeText & "」" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "「" & Me.txtSearch.Text & "」をすべて「" & Me.txtReplace.Text & "」に置換します。" & vbCrLf
    confirmMsg = confirmMsg & "この操作は元に戻せません。よろしいですか？"
    
    If MsgBox(confirmMsg, vbYesNo + vbExclamation, "すべての置換の確認") = vbNo Then
        Exit Sub
    End If

    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Then Exit Sub

    ' --- 検索がまだ実行されていない場合は、まず検索を実行 ---
    If foundShapes Is Nothing Or lastSearchTerm <> searchTerm Then
        Call ExecuteSearch
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
            
            If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
                newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
                shp.TextFrame.Characters.Text = newText
                replacedCount = replacedCount + 1
            End If
        Next shp
        On Error GoTo 0

        MsgBox replacedCount & "個の項目を置換しました。", vbInformation
        
        Set foundShapes = Nothing
        lastSearchTerm = ""
        currentShapeIndex = 0

    Else
        MsgBox "置換対象の図形が見つかりませんでした。", vbExclamation
    End If
End Sub


' 検索範囲のオプションが変更されたら、検索結果をリセットする
Private Sub optAllSheets_Click()
    Set foundShapes = Nothing
End Sub

Private Sub optCurrentSheet_Click()
    Set foundShapes = Nothing
End Sub


' 「終了」ボタンが押されたときの処理
Private Sub btnClose_Click()
    Unload Me
End Sub


' フォームが閉じられるときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundShapes = Nothing
End Sub
