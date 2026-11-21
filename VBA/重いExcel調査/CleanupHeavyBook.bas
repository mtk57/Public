Attribute VB_Name = "CleanupHeavyBook"
Option Explicit

' 選んだブック(本番ファイルのコピー)から、ユーザー定義のセルスタイルを削除して上書き保存する。
Public Sub DeleteCustomStyles()
    Dim path As Variant, wb As Workbook, st As Style
    Dim i As Long, before As Long, after As Long, deleted As Long, failed As Long
    Dim started As Double, prevCalc As XlCalculation, prevSec As MsoAutomationSecurity

    path = Application.GetOpenFilename("Excelブック (*.xls*),*.xls*", , "本番ファイルのコピーを選択")
    If VarType(path) = vbBoolean Then Exit Sub
    If MsgBox("次のファイルのユーザー定義スタイルを削除して、上書き保存します。" & vbCrLf & _
              "必ずコピーのほうを選んでください。" & vbCrLf & vbCrLf & path, _
              vbOKCancel + vbExclamation, "スタイルの削除") <> vbOK Then Exit Sub

    started = Timer
    prevCalc = Application.Calculation
    prevSec = Application.AutomationSecurity
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Application.StatusBar = "ファイルを開いています..."
    Set wb = Workbooks.Open(Filename:=CStr(path), UpdateLinks:=0)
    before = wb.Styles.Count

    For i = wb.Styles.Count To 1 Step -1
        Set st = wb.Styles(i)
        If Not st.BuiltIn Then
            On Error Resume Next
            st.Delete
            If Err.Number = 0 Then deleted = deleted + 1 Else failed = failed + 1
            Err.Clear
            On Error GoTo Fail
        End If
        If i Mod 500 = 0 Then
            Application.StatusBar = "スタイルを削除中... 残り " & i & " / " & before
            DoEvents
        End If
    Next i

    after = wb.Styles.Count
    Application.StatusBar = "保存しています..."
    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing

    RestoreSettings prevCalc, prevSec
    MsgBox "スタイルの削除が完了しました。" & vbCrLf & vbCrLf & _
           "削除前: " & before & vbCrLf & _
           "削除後: " & after & vbCrLf & _
           "削除した数: " & deleted & vbCrLf & _
           "削除できなかった数: " & failed & vbCrLf & _
           "所要時間(秒): " & Format(ElapsedSeconds(started), "0.0"), vbInformation, "スタイルの削除"
    Exit Sub

Fail:
    Dim msg As String
    msg = "エラーで中断しました。ファイルは保存していません。" & vbCrLf & _
          Err.Number & " / " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    RestoreSettings prevCalc, prevSec
    MsgBox msg, vbCritical, "スタイルの削除"
End Sub

' 変更したExcelの設定を元に戻す。
Private Sub RestoreSettings(ByVal prevCalc As XlCalculation, ByVal prevSec As MsoAutomationSecurity)
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' 日付をまたいでも負にならない経過秒数。
Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
