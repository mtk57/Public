VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private dst_not_found_action As String
Private is_copy_format As Boolean
Private is_copy_format_s As String
Private is_backup As Boolean
Private is_backup_s As String
Private is_not_close As Boolean
Private is_not_close_s As String
Private is_skip_blank As Boolean
Private is_skip_blank_s As String
Private is_ignore_case As Boolean
Private is_ignore_case_s As String

'コンストラクタ
Private Sub Class_Initialize()
    dst_not_found_action = ""
    is_copy_format = False
    is_backup = False
    is_not_close = False
    is_skip_blank = False
    is_ignore_case = True
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const Clm = "G"
    Const i = 7
    
    dst_not_found_action = main_sheet.Range(Clm & i).value
    is_copy_format_s = main_sheet.Range(Clm & i + 1).value
    is_backup_s = main_sheet.Range(Clm & i + 2).value
    is_not_close_s = main_sheet.Range(Clm & i + 3).value
    is_skip_blank_s = main_sheet.Range(Clm & i + 4).value
    is_ignore_case_s = main_sheet.Range(Clm & i + 5).value
End Sub

'検証
Public Sub Validate()
    If dst_not_found_action = "" Or _
       dst_not_found_action = "中断" Then
       dst_not_found_action = "中断"
    Else
        dst_not_found_action = "無視"
    End If

    If is_copy_format_s = "" Or _
       is_copy_format_s = "NO" Then
       is_copy_format = False
    Else
        is_copy_format = True
    End If

    If is_backup_s = "" Or _
       is_backup_s = "NO" Then
       is_backup = False
    Else
        is_backup = True
    End If

    If is_not_close_s = "" Or _
       is_not_close_s = "NO" Then
       is_not_close = False
    Else
        is_not_close = True
    End If
    
    If is_skip_blank_s = "" Or _
       is_skip_blank_s = "NO" Then
       is_skip_blank = False
    Else
        is_skip_blank = True
    End If

    If is_ignore_case_s = "" Or _
       is_ignore_case_s = "YES" Then
       is_ignore_case = True
    Else
        is_ignore_case = False
    End If
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'DST検索列が見つからない場合
'-----------------------------------------
Public Function GetDestNotFoundAction() As String
    GetDestNotFoundAction = dst_not_found_action
End Function

'-----------------------------------------
'書式もコピーするか?
'-----------------------------------------
Public Function IsCopyFormat() As Boolean
    IsCopyFormat = is_copy_format
End Function

'-----------------------------------------
'バックアップ作成
'-----------------------------------------
Public Function IsBackup() As Boolean
    IsBackup = is_backup
End Function

'-----------------------------------------
'転記後にファイルを閉じない
'-----------------------------------------
Public Function IsNotClose() As Boolean
    IsNotClose = is_not_close
End Function

'-----------------------------------------
'転記元が空の場合は無視する
'-----------------------------------------
Public Function IsSkipBlank() As Boolean
    IsSkipBlank = is_skip_blank
End Function

'-----------------------------------------
'検索時に大文字小文字を区別する
' True=区別する, False=区別しない
'-----------------------------------------
Public Function IsIgnoreCase() As Boolean
    IsIgnoreCase = is_ignore_case
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "dst_not_found_action=(" & dst_not_found_action & "), " & _
          "is_copy_format=(" & is_copy_format_s & "), " & _
          "is_backup=(" & is_backup_s & "), " & _
          "is_not_close=(" & is_not_close_s & "), " & _
          "is_skip_blank=(" & is_skip_blank_s & "), " & _
          "is_ignore_case=(" & is_ignore_case_s & ")"

    GetAllValue = ret
End Function
