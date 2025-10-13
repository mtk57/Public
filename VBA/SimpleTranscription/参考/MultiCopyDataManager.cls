VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiCopyDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数
Private sub_param As subparam
Private copy_datas() As CopyData
Private keyword_list() As String

'コンストラクタ
Private Sub Class_Initialize()

End Sub

'初期化
Public Sub Init(ByRef subprm As subparam, ByRef copydatas() As CopyData)
    Set sub_param = subprm
    copy_datas = copydatas
    
    CreateKeywordList
End Sub

'キーワードリスト(昇順、重複なし)を返す
Public Function GetKeywordList() As String()
    GetKeywordList = keyword_list
End Function

'キーワードリストから指定キーワードの数をカウントして返す
Public Function GetKeywordCount(ByVal keyword As String) As Long
    Dim ret As Long: ret = 0
    
    If keyword = "" Then
        GetKeywordCount = ret
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(copy_datas)
        If copy_datas(i).GetKey() = keyword Then
            ret = ret + 1
        End If
    Next i
    
    GetKeywordCount = ret
End Function

'指定キーワードからコピー値リストを作成して返す
Public Function GetValues(ByVal keyword As String) As String()
    Dim ret() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    
    For i = 0 To UBound(copy_datas)
        If copy_datas(i).GetKey() = keyword Then
            ReDim Preserve ret(cnt)
            ret(cnt) = copy_datas(i).GetValue()
            cnt = cnt + 1
        End If
    Next i
    
    GetValues = ret
End Function

Public Sub SetIsInserted(ByVal is_inserted As Boolean)
    sub_param.SetIsInserted (is_inserted)
End Sub

'キーワードリスト(昇順、重複なし)を作成する
Private Sub CreateKeywordList()
    Dim i As Long
    
    For i = 0 To UBound(copy_datas)
        ReDim Preserve keyword_list(i)
        keyword_list(i) = copy_datas(i).GetKey()
    Next i
    
    keyword_list = Common.SortAndDistinctArray(keyword_list)
End Sub


