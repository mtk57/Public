VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private is_enable As String
Private row_number As Long

'SRC
Private src_file_path As String
Private src_sheet_name As String
Private src_find_clm As String
Private src_tran_clm As String
Private src_ext As String

'DST
Private dst_file_path As String
Private dst_sheet_name As String
Private dst_find_clm As String
Private dst_tran_clm As String
Private dst_ext As String


'コンストラクタ
Private Sub Class_Initialize()
    is_enable = "STOPPER"
    row_number = 1
    
    src_file_path = ""
    src_sheet_name = ""
    src_find_clm = ""
    src_tran_clm = ""
    src_ext = ""
    
    dst_file_path = ""
    dst_sheet_name = ""
    dst_find_clm = ""
    dst_tran_clm = ""
    dst_ext = ""
End Sub

'初期化
Public Sub Init(ByVal row_num As Long)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Dim i As Long: i = row_num
    row_number = row_num
    
    is_enable = main_sheet.Range("D" & i).value
    
    src_file_path = main_sheet.Range("E" & i).value
    src_sheet_name = main_sheet.Range("F" & i).value
    src_find_clm = main_sheet.Range("G" & i).value
    src_tran_clm = main_sheet.Range("H" & i).value
    
    dst_file_path = main_sheet.Range("I" & i).value
    dst_sheet_name = main_sheet.Range("J" & i).value
    dst_find_clm = main_sheet.Range("K" & i).value
    dst_tran_clm = main_sheet.Range("L" & i).value
End Sub

'検証
Public Sub Validate()
    Dim SEP As String: SEP = Application.PathSeparator

    If is_enable = "" Or _
       is_enable = "STOPPER" Then
       is_enable = "STOPPER"
    ElseIf is_enable = "ENABLE" Then
        is_enable = "ENABLE"
    Else
        is_enable = "DISABLE"
    End If

    If is_enable = "DISABLE" Or is_enable = "STOPPER" Then
        Exit Sub
    End If
    
    If src_file_path = "" Then
        Err.Raise 53, , "SRCファイルパスが指定されていません"
    End If
    
    If InStr(src_file_path, SEP) > 0 Then
        If Common.IsExistsFile(src_file_path) = False Then
            Err.Raise 53, , "SRCファイルが存在しません" & vbCrLf & _
                            src_file_path
        End If
    Else
        If Common.IsOpenWorkbook(src_file_path) = False Then
            Err.Raise 53, , "SRCファイルが開かれていません" & vbCrLf & _
                            src_file_path
        End If
    End If
    
    src_ext = Common.GetFileExtension(src_file_path)

    If src_sheet_name = "" Then
        Err.Raise 53, , "SRCシート名が指定されていません"
    End If

    If src_find_clm = "" Then
        Err.Raise 53, , "SRC検索列が指定されていません"
    Else
        'TODO:列名の妥当性チェック
    End If
    
    If src_tran_clm = "" Then
        Err.Raise 53, , "SRC転記列が指定されていません"
    Else
        'TODO:列名の妥当性チェック
    End If
    
    If dst_file_path = "" Then
        Err.Raise 53, , "DSTファイルパスが指定されていません"
    End If
    
    If InStr(dst_file_path, SEP) > 0 Then
        If Common.IsExistsFile(dst_file_path) = False Then
            Err.Raise 53, , "DSTファイルが存在しません" & vbCrLf & _
                            dst_file_path
        End If
    Else
        If Common.IsOpenWorkbook(dst_file_path) = False Then
            Err.Raise 53, , "DSTファイルが開かれていません" & vbCrLf & _
                            dst_file_path
        End If
    End If
    
    dst_ext = Common.GetFileExtension(dst_file_path)
    
    'CheckExtensionCombination

    If dst_sheet_name = "" Then
        dst_sheet_name = src_sheet_name
    End If

    If dst_find_clm = "" Then
        dst_find_clm = src_find_clm
    Else
        'TODO:列名の妥当性チェック
    End If
    
    If dst_tran_clm = "" Then
        dst_tran_clm = src_tran_clm
    Else
        'TODO:列名の妥当性チェック
    End If

End Sub

Private Sub CheckExtensionCombination()
    Dim isOK As Boolean: isOK = True
    If src_ext = "xls" Or src_ext = "xlsx" Or src_ext = "xlsm" Then
        If dst_ext = "csv" Or dst_ext = "tsv" Then
            isOK = False
        End If
    ElseIf src_ext = "csv" Or src_ext = "tsv" Then
        If dst_ext = "xls" Or dst_ext = "xlsx" Or dst_ext = "xlsm" Then
            isOK = False
        End If
    End If
    
    If isOK = False Then
        Err.Raise 53, , "SRCファイルパスとDSTファイルパスの拡張子の組み合わせが不正です" & vbCrLf & _
                        "src_ext= " & src_ext & vbCrLf & _
                        "dst_ext= " & dst_ext
    End If
End Sub


'=========================================
'Getter
'=========================================

'-----------------------------------------
'有効無効
' Ret : "ENABLE" or "DISABLE" or "STOPPER"
'-----------------------------------------
Public Function GetEnable() As String
    GetEnable = is_enable
End Function

'-----------------------------------------
'SubParamの行番号
'-----------------------------------------
Public Function GetSubParamRowNumber() As Long
    GetSubParamRowNumber = row_number
End Function

'-----------------------------------------
'SRCファイルパス
'-----------------------------------------
Public Function GetSrcFilePath() As String
    GetSrcFilePath = src_file_path
End Function

'-----------------------------------------
'SRCシート名
'-----------------------------------------
Public Function GetSrcSheetName() As String
    GetSrcSheetName = src_sheet_name
End Function

'-----------------------------------------
'SRC検索列
'-----------------------------------------
Public Function GetSrcFindClm() As String
    GetSrcFindClm = src_find_clm
End Function

'-----------------------------------------
'SRC転記列
'-----------------------------------------
Public Function GetSrcTranClm() As String
    GetSrcTranClm = src_tran_clm
End Function

'-----------------------------------------
'SRC拡張子
'-----------------------------------------
Public Function GetSrcExt() As String
    GetSrcExt = src_ext
End Function

'-----------------------------------------
'DSTファイルパス
'-----------------------------------------
Public Function GetDstFilePath() As String
    GetDstFilePath = dst_file_path
End Function

'-----------------------------------------
'DSTシート名
'-----------------------------------------
Public Function GetDstSheetName() As String
    GetDstSheetName = dst_sheet_name
End Function

'-----------------------------------------
'DST検索列
'-----------------------------------------
Public Function GetDstFindClm() As String
    GetDstFindClm = dst_find_clm
End Function

'-----------------------------------------
'DST転記列
'-----------------------------------------
Public Function GetDstTranClm() As String
    GetDstTranClm = dst_tran_clm
End Function

'-----------------------------------------
'DST拡張子
'-----------------------------------------
Public Function GetDstExt() As String
    GetDstExt = dst_ext
End Function


Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Sub Param] " & _
          "is_enable=(" & is_enable & "), " & _
          "src_file_path=(" & src_file_path & "), " & _
          "src_sheet_name=(" & src_sheet_name & "), " & _
          "src_find_clm=(" & src_find_clm & "), " & _
          "src_tran_clm=(" & src_tran_clm & "), " & _
          "dst_file_path=(" & dst_file_path & "), " & _
          "dst_sheet_name=(" & dst_sheet_name & "), " & _
          "dst_find_clm=(" & dst_find_clm & "), " & _
          "dst_tran_clm=(" & dst_tran_clm & ")"
    GetAllValue = ret
End Function

