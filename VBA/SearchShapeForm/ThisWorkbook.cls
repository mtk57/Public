VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --- ThisWorkbook の新しいコード ---

Private Const ShortcutKey As String = "^+s" ' Ctrl + Shift + S
Private Const MacroToRun As String = "ShowSearchShapeForm"

' アドインがインストール（有効化）されたときに実行
Private Sub Workbook_AddinInstall()
    Application.OnKey ShortcutKey, MacroToRun
End Sub

' アドインがアンインストール（無効化）されたときに実行
Private Sub Workbook_AddinUninstall()
    Application.OnKey ShortcutKey
End Sub

' Excel起動時にも念のため設定（ユーザーがアドインフォルダから直接開いた場合などを想定）
Private Sub Workbook_Open()
    Application.OnKey ShortcutKey, MacroToRun
End Sub
