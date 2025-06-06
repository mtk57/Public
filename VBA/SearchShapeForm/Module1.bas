Attribute VB_Name = "Module1"
' ショートカットキーからこのマクロを呼び出す
Public Sub ShowSearchShapeForm()
    ' フォームが既に開かれているかチェック
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "SearchShapeForm" Then
            ' 既に開いていれば何もしないで終了
            Exit Sub
        End If
    Next frm

    ' フォームをモードレスで表示
    SearchShapeForm.Show vbModeless
End Sub
