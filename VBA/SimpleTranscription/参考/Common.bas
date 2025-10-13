Attribute VB_Name = "Common"
Option Explicit

Private Const VERSION = "1.5.21"

Public Const REG_EX_VB_METHOD = "(Function|Sub)\s+[^\(\)\s]+\("
Public Const REG_EX_VB_METHOD_WITH_RET = "Function\s+[^\(\)\s]*\(.*\)(\s+As\s+[^\(\)\s]*\(*\)*)*$"


Public Type MethodInfoStruct
    Raw As String
    Name As String
    ret As String
    params() As String
End Type

Public Type GrepResultInfoStruct
    ResultRaw As String
    FilePath As String
    Ext As String
    row As Long
    Clm As Long
    Contents As String
    IsError As Boolean
    ErrorInfo As String
    MethodInfo As MethodInfoStruct
End Type

Public Enum GrepAppEnum
    sakura
    hidemaru
End Enum

Public Enum LangEnum
    VB6
    VBNET
    CSharp
End Enum

Private Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

Private Declare PtrSafe Sub GetLocalTime Lib _
    "kernel32" ( _
    lpSystemTime As SYSTEMTIME _
)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'ログファイル番号
Private logfile_num As Integer
Private is_log_opened As Boolean

Private Const GIT_BASH = "C:\Program Files\Git\usr\bin\bash.exe"

'-------------------------------------------------------------
' 指定されたExcelシート上の2つのセルの内容を比較し、違いがある部分を
' 視覚的にハイライトします。
'
' 引数:
' - sheetName (String): 比較を行うシートの名前
' - cellA_Ref (String): 比較元セルの参照（例: "A1"）
' - cellB_Ref (String): 比較先セルの参照（例: "B1"）
'
' 動作:
' 1. 指定されたシートをアクティブにします。
' 2. 比較元と比較先のセルの内容を取得します。
' 3. 2つの文字列を1文字ずつ比較します。
' 4. 比較先（cellB）で異なる部分を赤色でハイライトします。
' 5. 比較元（cellA）の内容は変更しません。
'
' 注意:
' - シートが存在しない場合、エラーメッセージを表示して処理を終了します。
' - 比較先の文字列が比較元よりも長い場合、余分な部分も赤くハイライトされます。
'-------------------------------------------------------------
Public Sub CompareCellsAndHighlight(ByVal sheetName As String, ByVal cellA_Ref As String, ByVal cellB_Ref As String)
    Dim ws As Worksheet
    Dim cellA As Range, cellB As Range
    Dim textA As String, textB As String
    Dim i As Long, j As Long, k As Long
    Dim diffStart As Long, diffLength As Long
    
    ' 指定されたシートを取得し、アクティブにする
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "指定されたシート '" & sheetName & "' が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ws.Activate
    
    ' 指定されたセルを取得
    Set cellA = ws.Range(cellA_Ref)
    Set cellB = ws.Range(cellB_Ref)
    
    ' セルの内容を取得
    textA = cellA.value
    textB = cellB.value
    
    ' セルBの書式をリセット
    cellB.Font.Color = RGB(0, 0, 0)
    
    i = 1
    j = 1
    
    Do While i <= Len(textA) Or j <= Len(textB)
        If i > Len(textA) Then
            ' textAの終わりに達した場合、残りのtextBを全てハイライト
            cellB.Characters(j, Len(textB) - j + 1).Font.Color = RGB(255, 0, 0)
            Exit Do
        ElseIf j > Len(textB) Then
            ' textBの終わりに達した場合、ループを終了
            Exit Do
        ElseIf Mid(textA, i, 1) = Mid(textB, j, 1) Then
            ' 文字が一致する場合
            i = i + 1
            j = j + 1
        Else
            ' 不一致を検出
            diffStart = j
            k = j
            
            ' 次の一致を探す
            Do While k <= Len(textB)
                If Mid(textA, i, 1) = Mid(textB, k, 1) Then
                    Exit Do
                End If
                k = k + 1
            Loop
            
            diffLength = k - j
            
            ' 不一致部分をハイライト
            cellB.Characters(diffStart, diffLength).Font.Color = RGB(255, 0, 0)
            
            j = k
            If k <= Len(textB) Then
                i = i + 1
                j = j + 1
            End If
        End If
    Loop
End Sub

'-------------------------------------------------------------
' VBの関数名を抽出する
'-------------------------------------------------------------
Public Function ExtractVBFunctionName(ByVal codeLine As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim functionName As String
    
    ' 正規表現オブジェクトを作成
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 関数定義を検出する正規表現パターン
    regex.Pattern = "(Private|Public|Protected)?\s*(Shared|MustOverride|Overridable|Overrides|Delegate|Overloads|Shadows|Static)?\s*(Function|Sub)\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\("
    regex.IgnoreCase = True
    regex.Global = False
    
    ' 正規表現でマッチングを実行
    Set matches = regex.Execute(codeLine)
    
    ' マッチした場合、関数名を抽出
    If matches.count > 0 Then
        functionName = matches(0).SubMatches(3)
    Else
        functionName = ""
    End If
    
    ExtractVBFunctionName = functionName
End Function

'-------------------------------------------------------------
' Function: FindWord
' 目的: 指定された条件に基づいて、ターゲット文字列内で検索文字列を見つけます。
'
' パラメータ:
'   targetStr (String) - 検索対象の文字列
'   findStr (String) - 検索する文字列またはパターン
'   letterCase (Boolean) - 大文字小文字を区別するかどうか
'                          True: 区別する、False: 区別しない
'   exactMatch (Boolean) - 完全一致で検索するかどうか
'                          True: 完全一致、False: 部分一致
'   useRegEx (Boolean) - 正規表現を使用するかどうか
'                        True: 使用する、False: 使用しない
'
' 戻り値:
'   Boolean - 検索文字列が見つかった場合はTrue、そうでない場合はFalse
'
' 注意:
'   1. 正規表現を使用する場合は、VBScriptの正規表現構文に従ってください。
'   2. 正規表現使用時は、exactMatchパラメータは無視されます。
'      完全一致を行いたい場合は、正規表現パターンで ^ と $ を使用してください。
'   3. 正規表現を使用する場合、「Microsoft VBScript Regular Expressions 5.5」
'      への参照設定が必要です。
'
' 使用例:
'   result = FindWord("Hello World", "world", False, False, False) ' 大小文字区別なし、部分一致
'   result = FindWord("Hello World", "^Hello World$", False, False, True) ' 正規表現による完全一致
'-------------------------------------------------------------
Public Function FindWord( _
    targetStr As String, _
    findStr As String, _
    Optional letterCase As Boolean = False, _
    Optional exactMatch As Boolean = False, _
    Optional useRegEx As Boolean = False _
) As Boolean
    
    Dim regex As Object
    
    If useRegEx Then
        ' 正規表現を使用する場合
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Pattern = findStr
            .IgnoreCase = Not letterCase
            .Global = True
            FindWord = .Test(targetStr)
        End With
    Else
        ' 正規表現を使用しない場合
        If exactMatch Then
            ' 完全一致の場合
            If letterCase Then
                FindWord = (targetStr = findStr)
            Else
                FindWord = (StrComp(targetStr, findStr, vbTextCompare) = 0)
            End If
        Else
            ' 部分一致の場合
            If letterCase Then
                FindWord = (InStr(1, targetStr, findStr, vbBinaryCompare) > 0)
            Else
                FindWord = (InStr(1, targetStr, findStr, vbTextCompare) > 0)
            End If
        End If
    End If
End Function

'-------------------------------------------------------------
' 最大値を返す
'-------------------------------------------------------------
Public Function Max(ByVal a As Variant, ByVal b As Variant) As Variant
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

'-------------------------------------------------------------
' 指定されたセルから列内の最後の使用済みセルまでをクリアする
' ws : I : ワークシート
' cell_address : I : セルアドレス(Ex."A1")
'-------------------------------------------------------------
Public Sub ClearRange(ByRef ws As Worksheet, ByVal cell_address As String)
    Dim last_row As Long
    Dim range_to_clear As Range

    ' 指定されたセルから列内の最後の使用済みセルまでの範囲を取得
    last_row = ws.Cells(ws.Rows.count, ws.Range(cell_address).Column).End(xlUp).row
    Set range_to_clear = ws.Range(cell_address, ws.Cells(last_row, ws.Range(cell_address).Column))

    ' 範囲をクリア
    'range_to_clear.Clear
    range_to_clear.value = ""
End Sub

'-------------------------------------------------------------
' 配列をクイックソートで昇順ソートする
' argAry : I/O : 配列
' lngMin : I : ソート範囲(最小)  ※任意
' lngMax : I : ソート範囲(最大)  ※任意
' keyPos : I : ソートキーインデクス  ※任意
' BASE : https://excel-ubara.com/excelvba5/EXCELVBA229.html
'-------------------------------------------------------------
Public Sub QuickSortArray( _
    ByRef argAry() As Variant, _
    Optional ByVal lngMin As Long = -1, _
    Optional ByVal lngMax As Long = -1, _
    Optional ByVal keyPos As Long = 0 _
)
    If IsEmptyArray(argAry) = True Then
        Exit Sub
    End If
    
    If lngMin = -1 Then
        lngMin = LBound(argAry)
    End If
    If lngMax = -1 Then
        lngMax = UBound(argAry)
    End If

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim vBase As Variant
    Dim vSwap As Variant
    
    vBase = argAry(Int((lngMin + lngMax) / 2), keyPos)
    i = lngMin
    j = lngMax
    
    Do
        Do While argAry(i, keyPos) < vBase
            i = i + 1
        Loop
        
        Do While argAry(j, keyPos) > vBase
            j = j - 1
        Loop
        
        If i >= j Then Exit Do
        
        For k = LBound(argAry, 2) To UBound(argAry, 2)
            vSwap = argAry(i, k)
            argAry(i, k) = argAry(j, k)
            argAry(j, k) = vSwap
        Next
        i = i + 1
        j = j - 1
    Loop
    
    If (lngMin < i - 1) Then
        Call QuickSortArray(argAry, lngMin, i - 1, keyPos)
    End If
    
    If (lngMax > j + 1) Then
        Call QuickSortArray(argAry, j + 1, lngMax, keyPos)
    End If
End Sub

'-------------------------------------------------------------
' 配列の指定位置に指定サイズの空行を挿入する
'  例:["A", "B", "C"]という配列の場合、以下のようになる
'    pos=2, size=2:["A", "B", "", "", "C"]
'    pos=0, size=1:["", "A", "B", "C"]
' ary : I : String配列
' pos : I : 挿入位置(0=先頭)
' size : I : 挿入行数
' RET : 挿入後のString配列
'-------------------------------------------------------------
Public Function InsertEmptyRows(ByRef ary() As String, ByVal pos As Long, ByVal size As Long) As String()
    '引数チェック
    If Common.IsEmptyArray(ary) = True Then
        InsertEmptyRows = ary
        Exit Function
    End If
    
    If UBound(ary) < pos Then
        Err.Raise 53, , "[InsertEmptyRows] posが配列サイズを超えています! (pos=" & pos & ")"
    End If
    
    If pos < 0 Then
        pos = 0
    End If
    
    If size < 1 Then
        InsertEmptyRows = ary
        Exit Function
    End If

    Dim i As Long
    Dim ret() As String
    ReDim ret(LBound(ary) To UBound(ary) + size)
    
    For i = LBound(ret) To UBound(ret)
        If i < pos Then
            ret(i) = ary(i)
        ElseIf i >= pos And i < pos + size Then
            ret(i) = ""
        Else
            ret(i) = ary(i - size)
        End If
    Next i
    
    InsertEmptyRows = ret
End Function

'-------------------------------------------------------------
' 指定した行の次の行から指定した行数分だけ行を追加する
' ws : I : ワークシート
' startRowNum : I : 指定行 (1～)
' insertRowCnt : I : 追加行数 (1～)
'-------------------------------------------------------------
Public Sub InsertRows(ByRef ws As Worksheet, ByVal startRowNum As Long, ByVal insertRowCnt As Long)
    '引数チェック
    If startRowNum < 1 Then
        startRowNum = 1
    End If
    
    If insertRowCnt < 1 Then
        Exit Sub
    End If
    
    ws.Rows(startRowNum + 1 & ":" & startRowNum + insertRowCnt).Insert shift:=xlDown
End Sub

'-------------------------------------------------------------
' Grep結果を解析してメソッド情報を返す
' grepResults : I :Grep結果の配列
'                   <sakuraの場合>
'                   ファイルフルパス + (行番号, 列番号) + " " + [エンコード] + ":" + 内容
' grepApp : I : Grepしたアプリ (現時点ではsakuraのみサポート)  ※任意
' lang : I : Grepした言語 (現時点ではVB6のみサポート)  ※任意
' Ret : メソッド情報 (Grep結果のコードが属しているメソッド)
'-------------------------------------------------------------
Public Function GetMethodInfoFromGrepResult( _
    ByRef grepResults() As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As GrepResultInfoStruct()

    Dim ret() As GrepResultInfoStruct

    '引数チェック
    If Common.IsEmptyArrayLong(grepResults) = True Then
         GetMethodInfoFromGrepResult = ret
         Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        GetMethodInfoFromGrepResult = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        GetMethodInfoFromGrepResult = ret
        Exit Function
    End If
    

    ReDim Preserve ret(UBound(grepResults))

    Dim i As Long
    Dim raw_contents() As String
    Dim before_path As String: before_path = ""
    
    For i = 0 To UBound(grepResults)
        Dim cur_info As GrepResultInfoStruct
        
        'Grep情報を取得する
        cur_info = GetGrepInfo(grepResults(i), grepApp, lang)
        
        'コメントの場合は無視する
        If IsCommentCode(cur_info.Contents, cur_info.Ext) Then
            GoTo CONTINUE_A
        End If
        
        If before_path = cur_info.FilePath Then
            '前回と同じファイルなのでメモリ上の内容を使いまわす
        Else
            '前回と異なるファイルに切り替わったのでファイルの内容を読み込む
            If IsExistsFile(cur_info.FilePath) = False Then
                'ファイルが存在しないので無視する
                cur_info.IsError = True
                cur_info.ErrorInfo = "★File not exist.(" & cur_info.FilePath & ")"
                GoTo CONTINUE_A
            End If
            
            raw_contents = GetContents(cur_info.FilePath)
            before_path = cur_info.FilePath
        End If
        

        '行番号の位置から上に向かってメソッドを探す
        cur_info.MethodInfo = FindMethodByGrepResultInfo(cur_info, raw_contents, grepApp, lang)
        
CONTINUE_A:

        ret(i) = cur_info
        
    Next i
    
    
    GetMethodInfoFromGrepResult = ret

End Function

'-------------------------------------------------------------
' Grep結果を解析してGREP情報を返す
' grepResult : I :Grep結果
'                   <sakuraの場合>
'                   ファイルフルパス + (行番号, 列番号) + " " + [エンコード] + ":" + 内容
' grepApp : I : Grepしたアプリ (現時点ではsakuraのみサポート)  ※任意
' lang : I : Grepした言語 (現時点ではVB6のみサポート)  ※任意
' Ret : GREP情報
'-------------------------------------------------------------
Public Function GetGrepInfo( _
    ByVal grepResult As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As GrepResultInfoStruct

    Dim ret As GrepResultInfoStruct

    '引数チェック
    If grepResult = "" Then
        GetGrepInfo = ret
        Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        GetGrepInfo = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        GetGrepInfo = ret
        Exit Function
    End If

    Dim regex_result() As String
    regex_result = GetMatchByRegExp(grepResult, "^.*\(\d+,\d+\)", False)
    
    If IsEmptyArray(regex_result) Then
        GetGrepInfo = ret
        Exit Function
    End If
    
    ret.ResultRaw = grepResult
    ret.FilePath = ReplaceByRegExp(regex_result(0), "\(\d+,\d+\)", "", False)
    ret.Ext = GetFileExtension(ret.FilePath)
    Dim wk As String: wk = Trim(Replace(Replace(GetMatchByRegExp(regex_result(0), "\(\d+,\d+\)", False)(0), "(", ""), ")", ""))
    ret.row = CLng(Split(wk, ",")(0))
    ret.Clm = CLng(Split(wk, ",")(1))
    ret.Contents = ReplaceByRegExp(grepResult, "^.*\]:", "", False)
    
    GetGrepInfo = ret
End Function

'-------------------------------------------------------------
'対象ファイルを読み込んで内容を配列で返す
' path : I : 対象ファイルパス
' Ret : 読み込んだ内容
'-------------------------------------------------------------
Public Function GetContents(ByVal path As String) As String()
    Dim ret() As String
    
    If IsExistsFile(path) = False Then
        GetContents = ret
        Exit Function
    End If
    
    If IsUTF8(path) = False Then
        GetContents = Split(ReadTextFileBySJIS(path), vbCrLf)
    Else
        GetContents = Split(Replace(ReadTextFileByUTF8(path), vbCrLf, vbLf), vbLf)
    End If

End Function

'-------------------------------------------------------------
' GREP結果情報からメソッドを探してメソッド情報として返す
' grepResult : I : GREP結果情報
' contents : I : 探す対象ファイルの内容
' grepApp : I : Grepしたアプリ (現時点ではsakuraのみサポート)  ※任意
' lang : I : Grepした言語 (現時点ではVB6のみサポート)  ※任意
' Ret : メソッド情報
'-------------------------------------------------------------
Public Function FindMethodByGrepResultInfo( _
    ByRef grepResult As GrepResultInfoStruct, _
    ByRef Contents() As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As MethodInfoStruct

    Dim ret As MethodInfoStruct
    
    '引数チェック
    If Common.IsEmptyArrayLong(Contents) = True Then
         FindMethodByGrepResultInfo = ret
         Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        FindMethodByGrepResultInfo = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        FindMethodByGrepResultInfo = ret
        Exit Function
    End If
    
    
    If lang = LangEnum.VB6 Or lang = VBNET Then
        Dim method_type As String
        Dim method_start_row As Long
    
        'メソッド(Function/Sub)のタイプを見つける (複数行は考慮しない)
        method_type = FindMethodTypeForVB(Contents, grepResult.row)
        
        If method_type = "" Then
            '発見できず。。
            FindMethodByGrepResultInfo = ret
            Exit Function
        End If
        
        'メソッド(Function/Sub)の開始行番号を見つける
        method_start_row = FindMethodStartRowForVB(Contents, grepResult.row, method_type)
        
        If method_start_row = -1 Then
            '発見できず。。
            FindMethodByGrepResultInfo = ret
            Exit Function
        End If
        
        'メソッド(Function/Sub)の開始行番号からメソッド情報を取得する
        ret = GetMethodInfoForVB(Contents, method_start_row, method_type)
        
    End If
    
    FindMethodByGrepResultInfo = ret

End Function

'-------------------------------------------------------------
' VBのメソッド(Function/Sub)の種類を見つける
' (複数行は考慮しない)
' contents : I : 探す対象ファイルの内容
' startRow : I : 開始行
' Ret : 発見(Function/Sub), 発見できず。("")
'-------------------------------------------------------------
Public Function FindMethodTypeForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long _
) As String

    Dim ret As String: ret = ""

    '引数チェック
    If Common.IsEmptyArrayLong(Contents) = True Then
         FindMethodTypeForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         FindMethodTypeForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    
    For i = 0 To UBound(Contents)
        If (startRow + i) > UBound(Contents) Then
            Exit For
        End If
        
        Dim line As String: line = Contents(startRow + i)
        
        If IsCommentCode(line, "bas") Then
            'コメント行は無視
            GoTo CONTINUE
        End If
        
        If IsMatchByRegExp(line, "End\s+(Function|Sub)\b", True) = False Then
            '見つからない
            GoTo CONTINUE
        End If
        
        '発見
        ret = Trim(Replace(line, "End", ""))
        Exit For
        
CONTINUE:
        
    Next i
    
    FindMethodTypeForVB = ret

End Function

'-------------------------------------------------------------
' VBのメソッド(Function/Sub)の開始行番号を見つける
' (引数、戻り値以外の複数行は考慮しない)
' contents : I : 探す対象ファイルの内容
' startRow : I : 開始行
' methodType : I : 種類(Function/Sub)
' Ret : 発見(開始行番号), 発見できず。(-1)
'-------------------------------------------------------------
Public Function FindMethodStartRowForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long, _
    ByVal methodType As String _
) As Long

    Dim ret As Long: ret = -1

    '引数チェック
    If Common.IsEmptyArrayLong(Contents) = True Then
         FindMethodStartRowForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         FindMethodStartRowForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    
    For i = 0 To UBound(Contents)
        If (startRow - i) < 0 Then
            Exit For
        End If
    
        Dim line As String: line = Contents(startRow - i)
        
        If IsCommentCode(line, "bas") Then
            'コメント行は無視
            GoTo CONTINUE
        End If
        
        If IsMatchByRegExp(line, REG_EX_VB_METHOD, True) = False Then
            '見つからない
            GoTo CONTINUE
        End If
        
        '発見
        ret = startRow - i
        Exit For
        
CONTINUE:
        
    Next i
    
    FindMethodStartRowForVB = ret

End Function

'-------------------------------------------------------------
' VBのメソッド(Function/Sub)情報を返す
' contents : I : 探す対象ファイルの内容
' startRow : I : メソッド開始行
' methodType : I : 種類(Function/Sub)
' Ret : メソッド情報
'-------------------------------------------------------------
Public Function GetMethodInfoForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long, _
    ByVal methodType As String _
) As MethodInfoStruct

    Dim ret As MethodInfoStruct

    '引数チェック
    If Common.IsEmptyArrayLong(Contents) = True Then
         GetMethodInfoForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         GetMethodInfoForVB = ret
         Exit Function
    End If
    
    If methodType <> "Function" And methodType <> "Sub" Then
         GetMethodInfoForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    Dim merge_lines As String
    
    Dim start_clm As Long
    Dim end_clm As Long
    
    'いったん複数行を1行にまとめる
    For i = 0 To UBound(Contents)
        If (startRow + i) > UBound(Contents) Then
            Exit For
        End If
    
        Dim line As String: line = Contents(startRow + i)
        
        If IsCommentCode(line, "bas") Then
            'コメント行は無視
            GoTo CONTINUE
        End If
        
        'コメント削除
        line = Trim(RemoveRightComment(line, "bas"))
        
        '"_"削除
        line = Trim(ReplaceByRegExp(line, "_$", "", False))
    
        'マージ
        merge_lines = merge_lines + line
        
        
        '引数の終わりの括弧を探す
        end_clm = FindMatchingBracketPositionForVB(merge_lines)
        
        If end_clm = 0 Then
            '見つからない
            GoTo CONTINUE
        End If
        
        '見つかった
        
        If methodType = "Sub" Then
            Exit For
        End If
        
        'Functionの場合は戻り値までマージされているか確認する
        If IsMatchByRegExp(merge_lines, REG_EX_VB_METHOD_WITH_RET, True) = True Then
            Exit For
        End If
        
CONTINUE:
        
    Next i
    
    
    '各情報を取得
    ret.Raw = merge_lines
    
    'メソッド名
    Dim wk As String: wk = GetMatchByRegExp(merge_lines, REG_EX_VB_METHOD, True)(0)
    ret.Name = Replace(Replace(wk, methodType & " ", ""), "(", "")
    
    start_clm = InStr(merge_lines, "(")
    end_clm = FindMatchingBracketPositionForVB(merge_lines)
    
    '戻り値
    If methodType = "Function" Then
        wk = Mid(merge_lines, end_clm)
        
        If wk = ")" Then
            '戻り値がないFunction
            ret.ret = ""
        Else
            ret.ret = Replace(wk, ") As ", "")
        End If
    Else
        ret.ret = ""
    End If
    
    '引数
    wk = Left(Mid(merge_lines, start_clm + 1), end_clm - start_clm - 1)
    ret.params = Split(wk, ",")
    
    For i = 0 To UBound(ret.params)
        ret.params(i) = Trim(ret.params(i))
    Next i
    
    GetMethodInfoForVB = ret

End Function

'-------------------------------------------------------------
' 最初に見つけた"("に対応する")"を探して桁位置を返す
' inputString : I : 対象文字列
' Ret : 0=見つからなかった, 1以上=")"の桁位置
'-------------------------------------------------------------
Public Function FindMatchingBracketPositionForVB(ByVal inputString As String) As Long

    Dim ret As Long: ret = 0

    '引数チェック
    If inputString = "" Or _
       InStr(inputString, "(") = 0 Or InStr(inputString, ")") = 0 Then
        FindMatchingBracketPositionForVB = ret
        Exit Function
    End If

    Dim openBracket As String
    Dim closeBracket As String
    Dim i As Long
    Dim stack As Long
    
    openBracket = "("
    closeBracket = ")"
    
    ' スタックを初期化
    stack = 0
    
    For i = 1 To Len(inputString)
        If Mid(inputString, i, 1) = openBracket Then
            ' 開始括弧
            stack = stack + 1
        ElseIf Mid(inputString, i, 1) = closeBracket Then
            ' 終了括弧
            stack = stack - 1
            If stack = 0 Then
                ' 対応する終了括弧が見つかった
                FindMatchingBracketPositionForVB = i
                Exit Function
            End If
        End If
    Next i
    
    '対応する終了括弧が見つからない
    FindMatchingBracketPositionForVB = ret
    
End Function

'-------------------------------------------------------------
' A1形式の文字列から列番号を返す
'  Ex. "A1" -> 1
'      "ZZ12" -> 27
' a1：I : A1形式の文字列
' Ret：列番号
'-------------------------------------------------------------
Public Function GetColNumFromA1(ByVal a1 As String) As Long
    Dim Clm As String
    Dim row As Long
    Call SplitCellAddress(a1, Clm, row)
    
    Dim substr As String
    substr = Clm
    
    Dim i As Integer
    
    For i = 1 To Len(Clm)
        If (i = 1) Then
            GetColNumFromA1 = GetColNumFromA1 + A_to_ColNum(Right(substr, 1))
        Else
            GetColNumFromA1 = GetColNumFromA1 + (A_to_ColNum(Right(substr, 1)) * (i - 1) * 26)
        End If
        substr = Left(substr, Len(substr) - 1)
    Next i
End Function

'-------------------------------------------------------------
' A～Zを1～26に変換する
' idx：I : A～Z
' Ret：1～26
'-------------------------------------------------------------
Public Function A_to_ColNum(ByVal az As String) As Integer
    A_to_ColNum = Asc(az) - 65 + 1
End Function

'-------------------------------------------------------------
' A1形式のセルアドレスを列名と行番号に分離して返す
' cell_adr : I : A1形式のセルアドレス
' clm_name : O : 列名
' row_num : O : 行番号
' Ret : True=Success, False=Failed
'-------------------------------------------------------------
Public Function SplitCellAddress(ByVal cell_adr As String, ByRef clm_name As String, ByRef row_num As Long) As Boolean
    Dim matches() As String
    
    matches = GetMatchByRegExp(cell_adr, "[A-Z]+", True)
    
    If Common.IsEmptyArray(matches) = True Then
         SplitCellAddress = False
         Exit Function
    End If
    
    clm_name = matches(0)
    
    matches = GetMatchByRegExp(cell_adr, "[0-9]+", False)
    
    If Common.IsEmptyArray(matches) = True Then
         SplitCellAddress = False
         Exit Function
    End If
    
    row_num = CLng(matches(0))
    
    SplitCellAddress = True
End Function

'-------------------------------------------------------------
' ファイルに文字列リストをUTF-8で書き込む
' path : I : 指定ファイルパス(絶対パス)
' str_ary : I : 文字列リスト
'-------------------------------------------------------------
Public Sub SaveToFileFromStringArray(ByVal path As String, ByRef str_ary() As String)
    If path = "" Or IsExistsFile(path) = False Then
        Err.Raise 53, , "[SaveToFileFromStringArray] 指定されたパスが不正です (path=" & path & ")"
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
On Error GoTo Error
    '文字コードをUTF-8に設定する
    stream.Charset = "UTF-8"
    
    'テキストモードで開く
    stream.Open
    
    Dim row As Long
    Dim line As String
    
    For row = 0 To UBound(str_ary)
        line = str_ary(row)
        stream.WriteText line
        stream.WriteText vbCrLf
    Next row
    
    Const OVER_WRITE = 2
    stream.SaveToFile path, OVER_WRITE

    stream.Close
    Set stream = Nothing
    
    Exit Sub
Error:
    stream.Close
    Set stream = Nothing
    
    Err.Raise 53, , "[SaveToFileFromStringArray] エラー! (path=" & path & "), Desc=" & Err.Description
End Sub

'-------------------------------------------------------------
'文字列を末尾から先頭に向かって見ていき、指定された文字を見つけたらそこまでの文字列を返す
' 例:str="ABC:DEF", last_char=":"の場合、"DEF"が返る
' str : I : 文字列
' last_char : I : 指定された文字(1文字)
' Ret : 指定された文字を見つけたらそこまでの文字列。見つからない場合は""
'-------------------------------------------------------------
Public Function GetStringLastChar(ByVal str As String, ByVal last_char As String) As String
    '文字列の長さを取得
    Dim length As Integer
    Dim i As Integer
    Dim ch As String
    
    length = Len(str)
    
    If length = 0 Then
        GetStringLastChar = ""
        Exit Function
    End If
    
    '文字列の末尾から先頭に向かってループ
    For i = length To 1 Step -1
        'i番目の文字を取得
        ch = Mid(str, i, 1)
        
        '見つかった
        If ch = last_char Then
            GetStringLastChar = Right(str, length - i)
            Exit Function
        End If
    Next i
    
    '見つからなかった
    GetStringLastChar = ""
End Function

'-------------------------------------------------------------
'パスが255byte以上かを返す
' path : I : パス (絶対・相対はチェックしない)
' Ret : True/False (True=255byte以上, False=255byte未満)
'-------------------------------------------------------------
Public Function IsMaxOverPath(ByVal path As String) As Boolean
    IsMaxOverPath = LenB(StrConv(path, vbFromUnicode)) >= 255
End Function

'-------------------------------------------------------------
'文字列が指定文字列で開始されているかを返す
' target : I : 文字列
' search : I : 指定文字列
' Ret : True/False (True=開始されている, False=開始されていない)
'-------------------------------------------------------------
Public Function StartsWith(ByVal target As String, ByVal search As String) As Boolean
    StartsWith = False
    
    If Len(search) > Len(target) Then
        Exit Function
    End If
    
    If Left(target, Len(search)) = search Then
        StartsWith = True
    End If
    
End Function

'-------------------------------------------------------------
'フォルダが空かどうかを返す
' path : I : フォルダパス(絶対パス)
' Ret : True/False (True=空, False=空では無い)
'-------------------------------------------------------------
Public Function IsEmptyFolder(ByVal path As String) As Boolean
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[IsEmptyFolder] 指定されたフォルダが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsEmptyFolder] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Dim folder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    IsEmptyFolder = folder.Files.count = 0 And folder.SubFolders.count = 0
    
    Set fso = Nothing
    Set folder = Nothing
End Function

'-------------------------------------------------------------
'String配列を昇順ソートして重複行を削除して返す
' arr : I : 配列
' Ret : 昇順ソートして重複行を削除した配列
'-------------------------------------------------------------
Public Function SortAndDistinctArray(ByRef arr() As String) As String()
    Dim dict_ As Object
    Set dict_ = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not dict_.Exists(arr(i)) Then
            dict_.Add arr(i), 1
        End If
    Next i
    Dim result() As String
    ReDim result(0 To dict_.count - 1)
    Dim key As Variant
    i = 0
    For Each key In dict_.Keys()
        result(i) = key
        i = i + 1
    Next key
    Set dict_ = Nothing
    SortAndDistinctArray = result
End Function

'-------------------------------------------------------------
'右のコメントを削除して返す
' str : I : 文字列
' ext : I : 拡張子(Ex. "bas", "vb") ※VB系のみサポート
' Ret : コメントがあれば削除して返す。なければ元の文字列を返す
' Ex. "abc 'def" → "abc"
'-------------------------------------------------------------
Public Function RemoveRightComment(ByVal str As String, ByVal Ext As String) As String
    Dim pos As Long
    Dim ret As String
    
    If Ext = "bas" Or _
       Ext = "frm" Or _
       Ext = "cls" Or _
       Ext = "ctl" Or _
       Ext = "vb" Then
        pos = InStr(str, "'")
        
        If pos = 0 Then
            ret = str
        Else
            ret = RTrim(Mid(str, 1, pos - 1))
        End If
    Else
        Err.Raise 53, , "[RemoveRightComment] 指定された拡張子は未サポートです (ext=" & Ext & ")"
    End If
    
    RemoveRightComment = RTrim(ret)

End Function

'-------------------------------------------------------------
'最初に見つかった英字の位置を返す
' str : I : 文字列
' Ret : 英字の位置(見つからない場合は0を返す)
'-------------------------------------------------------------
Public Function FindFirstCasePosition(ByVal str As String) As Long
    Dim i As Long
    Dim char As String
    
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If char Like "[A-Za-z]" Then
            FindFirstCasePosition = i
            Exit Function
        End If
    Next i
    
    FindFirstCasePosition = 0
End Function

'-------------------------------------------------------------
'コメント行かを判定する
' line : I : 行データ
' ext : I : 拡張子(Ex. "bas", "vb") ※VB系のみサポート
' Ret : True/False(True=コメント行)
'-------------------------------------------------------------
Public Function IsCommentCode(ByVal line As String, ByVal Ext As String) As Boolean
    If line = "" Or Ext = "" Then
        IsCommentCode = False
        Exit Function
    End If
    
    Dim wk As String
    wk = Replace(line, vbTab, " ")
    
    If Ext = "bas" Or _
       Ext = "frm" Or _
       Ext = "cls" Or _
       Ext = "ctl" Or _
       Ext = "vb" Then
        If Left(LTrim(wk), 1) = "'" Or _
           Left(LTrim(wk), 4) = "REM " Then
           IsCommentCode = True
           Exit Function
        End If
    Else
        Err.Raise 53, , "[IsCommentCode] 指定された拡張子は未サポートです (ext=" & Ext & ")"
    End If
    
    IsCommentCode = False

End Function

'-------------------------------------------------------------
'フォルダパスに指定フォルダ名があるかチェックし、あればそのフォルダまでのパスを返す
' path : I : フォルダパス(絶対パス)
' keyword : I : キーワード
' Ret : キーワードまでのパス(キーワードが見つからない場合は空を返す)
'-------------------------------------------------------------
Public Function GetFolderPathByKeyword( _
    path As String, _
    keyword As String _
) As String
    If path = "" Or keyword = "" Then
        GetFolderPathByKeyword = ""
        Exit Function
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderPathByKeyword] パスが長すぎます (path=" & path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim path_ary() As String
    Dim ret_ary() As String
    Dim i As Integer
    Dim j As Integer
    
    path_ary = Split(path, SEP)

    For i = UBound(path_ary) To 0 Step -1
        If path_ary(i) = keyword Then
        
            ReDim Preserve ret_ary(i)
            
            For j = LBound(ret_ary) To UBound(ret_ary)
                ret_ary(j) = path_ary(j)
            Next j
        
            GetFolderPathByKeyword = Join(ret_ary, SEP)
            Exit Function
        End If
    Next i
    
    GetFolderPathByKeyword = ""
End Function

'-------------------------------------------------------------
'フォルダパスから最後のフォルダ名を返す
' path : I : フォルダパス(絶対パス)
' Ret : 最後のフォルダ名
'        例: "C:\abc\def\xyz" → "xyz"
'-------------------------------------------------------------
Public Function GetLastFolderName(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetLastFolderName] パスが長すぎます (path=" & path & ")"
    End If

    Dim new_path As String: new_path = Common.RemoveTrailingBackslash(path)
    GetLastFolderName = Right(new_path, Len(new_path) - InStrRev(new_path, Application.PathSeparator))
End Function

'-------------------------------------------------------------
'フォルダパスの末尾に現在日時の文字列を付与して返す
' path : I : フォルダパス(絶対パス)
' Ret : 末尾に現在日時の文字列を付与したファイルパス
'-------------------------------------------------------------
Public Function ChangeUniqueDirPath(ByVal path As String) As String
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[ChangeUniqueDirPath] 指定されたフォルダが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ChangeUniqueDirPath] パスが長すぎます (path=" & path & ")"
    End If

    Dim new_path As String
    new_path = path & "_" & GetNowTimeString()
    If IsExistsFolder(new_path) = True Then
        WaitSec 1
        new_path = path & "_" & GetNowTimeString()
    End If

    ChangeUniqueDirPath = new_path
End Function

'-------------------------------------------------------------
'正規表現でパターンマッチングした文字列を置換する
' test_str : I : 対象文字列
' ptn : I : 検索パターン
' replace_str : I : 置換後文字列
' is_ignore_case : I : 大文字小文字を区別するか(True=する)
' Ret : 置換後の対象文字列
' Note:
'  - 参照設定に以下を追加する
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function ReplaceByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal replace_str As String, _
    ByVal is_ignore_case As Boolean _
) As String
    Dim REG As New VBScript_RegExp_55.RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim list() As String
    ReDim list(0)
    
    REG.Global = True
    REG.IgnoreCase = is_ignore_case
    REG.Pattern = ptn
    
    ReplaceByRegExp = REG.Replace(test_str, replace_str)

End Function

'-------------------------------------------------------------
'正規表現でパターンマッチングした結果を返す
' test_str : I : 対象文字列
' ptn : I : 検索パターン
' is_ignore_case : I : 大文字小文字を区別するか(True=する)
' Ret : マッチした文字列リスト
' Note:
'  - 参照設定に以下を追加する
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function GetMatchByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal is_ignore_case As Boolean _
) As String()
    Dim REG As New VBScript_RegExp_55.RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim list() As String
    ReDim list(0)
    
    REG.Global = True
    REG.IgnoreCase = is_ignore_case
    REG.Pattern = ptn
    
    Set mc = REG.Execute(test_str)
    For Each m In mc
        Common.AppendArray list, m.value
    Next
    
    list = Common.DeleteEmptyArray(list)
    
    GetMatchByRegExp = list
End Function

'-------------------------------------------------------------
'正規表現でパターンマッチングを行う
' test_str : I : 対象文字列
' ptn : I : 検索パターン
' is_ignore_case : I : 大文字小文字を区別するか(True=する)
' Ret : True/False (True=一致)
' Note:
'  - 参照設定に以下を追加する
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function IsMatchByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal is_ignore_case As Boolean _
) As Boolean
    Dim REG As New VBScript_RegExp_55.RegExp
    REG.Global = True
    REG.IgnoreCase = is_ignore_case
    REG.Pattern = ptn
    
    IsMatchByRegExp = REG.Test(test_str)
End Function

'-------------------------------------------------------------
'自身のフォルダパスを返す
' Ret : フォルダパス
'-------------------------------------------------------------
Public Function GetMyDir() As String
    Dim currentProject As Workbook
    Set currentProject = ThisWorkbook
    GetMyDir = currentProject.path
End Function

'-------------------------------------------------------------
'文字列配列を連結して文字列を返す
' ary : I : 文字列配列
' delim : I : 区切り文字(1文字)
' with_dbl_quot : I : ダブルクォーテーションで囲むか否か (True=囲む)
' Ret : 区切り文字で連結後の文字列
'-------------------------------------------------------------
Public Function JoinFromArray(ByRef ary() As String, ByVal delim As String, ByVal with_dbl_quot As Boolean) As String
    If IsEmptyArray(ary) = True Or delim = "" Then
        JoinFromArray = ""
        Exit Function
    End If

    Dim ret As String: ret = ""
    Dim i As Long
    
    For i = LBound(ary) To UBound(ary)
        If with_dbl_quot = True Then
            ret = ret & Chr(34) & ary(i) & Chr(34) & delim
        Else
            ret = ret & ary(i) & delim
        End If
    Next i
    
    JoinFromArray = Left(ret, Len(ret) - 1)

End Function

'-------------------------------------------------------------
'ブックが開いているか否かを返す
' book_name : I : ブック名
' Ret : True/False (True=開いている)
'-------------------------------------------------------------
Function IsOpenWorkbook(ByVal book_name As String) As Boolean
    Dim wb As Workbook
    Dim is_err As Boolean
    is_err = False

On Error Resume Next
    Set wb = Workbooks(book_name)
    
    If Err.Number <> 0 Then
        is_err = True
        Err.Clear
    End If

On Error GoTo 0
    If is_err = True Then
        IsOpenWorkbook = False
    Else
        IsOpenWorkbook = True
    End If
End Function

'-------------------------------------------------------------
'空ファイルか否かを返す
' path : I : ファイルパス(絶対パス)
' Ret : True/False (True=空ファイル)
'-------------------------------------------------------------
Public Function IsEmptyFile(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsEmptyFile] 指定されたファイルが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsEmptyFile] パスが長すぎます (path=" & path & ")"
    End If

    IsEmptyFile = (FileLen(path) = 0)
End Function

'-------------------------------------------------------------
'Variant型の配列をString型の配列に変換する
' arr : I : variant型の配列
' Ret : String型の配列
'-------------------------------------------------------------
Public Function VariantToStringArray(arr As Variant) As String()
    Dim ret_arr() As String
    Dim i As Long
    
    ReDim ret_arr(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        ret_arr(i) = CStr(arr(i))
    Next i
    
    VariantToStringArray = ret_arr
End Function

'-------------------------------------------------------------
'ファイル内のキーワードを含む行を削除して上書き保存する
' path : I : ファイルパス(絶対パス)
' keyword : I : キーワード
'-------------------------------------------------------------
Public Sub RemoveLinesWithKeyword(ByVal path As String, ByVal keyword As String)
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[RemoveLinesWithKeyword] 指定されたファイルが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RemoveLinesWithKeyword] パスが長すぎます (path=" & path & ")"
    End If

    If keyword = "" Then
        Exit Sub
    End If
    
    Dim fso As Object
    Dim file As Object
    Dim temp_file As Object
    Dim line As String
    Dim temp_ext As String: temp_ext = "." & GetNowTimeString()
    
    Const READ_ONLY = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(path, READ_ONLY)
    Set temp_file = fso.CreateTextFile(path & temp_ext, True)
    
    Do While Not file.AtEndOfStream
        line = file.ReadLine
        
        If InStr(line, keyword) = 0 Then
            temp_file.WriteLine line
        End If
    Loop
    
    file.Close
    temp_file.Close
    
    fso.DeleteFile path
    fso.MoveFile path & temp_ext, path
    
    Set temp_file = Nothing
    Set file = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'文字列からキーワードで検索し、ヒットしたキーワードから最後までの文字列を返す
' target : I : 検索対象の文字列
' keyword : I : 検索キーワード
' Ret : ヒットしたキーワードから最後までの文字列(見つからない場合は"")
' Ex.
'   target:"C:\abc\def\xyz\123.txt"
'   keyword:"def"
'   Ret:"def\xyz\123.txt"
'-------------------------------------------------------------
Function GetStringByKeyword(ByVal target As String, ByVal keyword As String) As String
    Dim pos As Long
    pos = InStr(target, keyword)
    If pos > 0 Then
        GetStringByKeyword = Mid(target, pos)
    Else
        GetStringByKeyword = ""
    End If
End Function

'-------------------------------------------------------------
'Gitコマンドを実行する
' repo_path : I : ローカルリポジトリフォルダパス(絶対パス)
' command : I : コマンド (Ex."git log --oneline")
' Ret : 標準出力
'-------------------------------------------------------------
Public Function RunGit(ByVal repo_path As String, ByVal command As String) As String()
    Dim err_msg As String: err_msg = ""
    Dim std_out() As String

    If IsMaxOverPath(repo_path) = True Then
        Err.Raise 53, , "[RunGit] パスが長すぎます (repo_path=" & repo_path & ")"
    End If

    If IsExistsFile(GIT_BASH) = False Then
        err_msg = "[RunGit] gitが見つかりません (" & GIT_BASH & ")"
        GoTo FINISH_3
    End If
    
    If IsExistsFolder(repo_path) = False Then
        If InStr(command, "git clone") = 0 Then
            err_msg = "[RunGit] 指定されたフォルダが存在しません (repo_path=" & repo_path & ")"
            GoTo FINISH_3
        End If
    End If
    
    'コマンド実行結果格納用の一時ファイルパス
    Dim temp As String: temp = GetTempFolder() & Application.PathSeparator & GetNowTimeString() & ".txt"

    'コマンド作成
    Dim run_cmd As String: run_cmd = GIT_BASH & _
                                     " --login -i -c & cd " & repo_path & " & " & _
                                     command & _
                                     " > " & temp & " 2>&1"
    WriteLog "[RunGit] run_cmd=" & run_cmd
    
    'コマンド実行
    Dim objShell As Object
    Dim objExec As Object
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd.exe /c " & Chr(34) & run_cmd & Chr(34))
    
    'プロセス完了時に通知を受け取る
    Do While objExec.Status = 0
        DoEvents
    Loop
    
    'プロセスの戻り値を取得する
    If objExec.ExitCode <> 0 Then
        err_msg = "[RunGit] プロセスの戻り値が0以外です (exit code=" & objExec.ExitCode & ")"
        
        If IsEmptyFile(temp) = True Then
            GoTo FINISH_2
        Else
            GoTo FINISH
        End If
        
    End If
    
    If IsEmptyFile(temp) = True Then
        GoTo FINISH_2
    End If
    
FINISH:
    If IsUTF8(temp) = False Then
        std_out = Split(ReadTextFileBySJIS(temp), vbCrLf)
    Else
        'std_out = Split(ReadTextFileByUTF8(temp), vbLf)
        std_out = Split(Replace(ReadTextFileByUTF8(temp), vbCrLf, vbLf), vbLf)
    End If

FINISH_2:
    DeleteFile (temp)
    
FINISH_3:
    Set objShell = Nothing
    Set objExec = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , err_msg & vbCrLf & "std_out=" & Join(std_out, ",")
    End If

    RunGit = std_out
End Function

'-------------------------------------------------------------
'一時フォルダパスを取得する
' Ret : 一時フォルダパス(絶対パス)
'-------------------------------------------------------------
Public Function GetTempFolder() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempFolder = fso.getSpecialFolder(2)
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイルをコピーする
' src_path : I : コピー元ファイルパス(絶対パス)
' dst_path : I : コピー先ファイルパス(絶対パス)
' is_create_dir : I : コピー先にフォルダを作成してコピーする(デフォルト=false)
'-------------------------------------------------------------
Public Sub CopyFile(ByVal src_path As String, ByVal dst_path As String, Optional ByVal is_create_dir As Boolean = False)
    If IsExistsFile(src_path) = False Then
        Err.Raise 53, , "[CopyFile] 指定されたファイルが存在しません (src_path=" & src_path & ")"
    End If

    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dst_path) = True Then
        Err.Raise 53, , "[CopyFile] パスが長すぎます (src_path=" & src_path & ", dst_path=" & dst_path & ")"
    End If

    If dst_path = "" Or src_path = dst_path Or IsExistsFile(dst_path) = True Then
        Exit Sub
    End If
    
    If is_create_dir = False Then
        FileCopy src_path, dst_path
        Exit Sub
    End If
    
    Dim dst_dir_path As String
    dst_dir_path = Common.GetFolderPath(dst_path)
    
    If Common.IsExistsFolder(dst_dir_path) = True Then
        FileCopy src_path, dst_path
        Exit Sub
    End If
    
    'コピー先フォルダパスが存在しないので作成する
    Common.CreateFolder (dst_dir_path)
    
    FileCopy src_path, dst_path
End Sub


'-------------------------------------------------------------
'フォルダをリネームする
' path : I : フォルダパス(絶対パス)
' rename : I : リネーム後のフォルダ名
' Ret : リネーム後のフォルダパス
'-------------------------------------------------------------
Public Function RenameFolder(ByVal path As String, ByVal rename As String) As String
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[RenameFolder] 指定されたフォルダが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RenameFolder] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(path)
    
    Dim err_msg As String
    Dim retry As Integer
    For retry = 0 To 3

On Error Resume Next
        folder.Name = rename
    
        err_msg = Err.Description
        Err.Clear
On Error GoTo 0

        If err_msg = "" Then
            Exit For
        End If
        
        WaitSec 1

    Next retry
    
    Set fso = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , "[RenameFolder] エラー! (err_msg=" & err_msg & ")"
    End If

    RenameFolder = folder.path

End Function

'-------------------------------------------------------------
'ワークシートの指定列のデータ最終行番号を返す
' ws : I : ワークシート
' clm : I : 指定列名(Ex."A")
'-------------------------------------------------------------
Public Function GetLastRowFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal Clm As String _
) As Long
    GetLastRowFromWorksheet = ws.Cells(ws.Rows.count, Clm).End(xlUp).row
End Function

'-------------------------------------------------------------
'文字列の配列から指定ワードで検索し、ヒットした行番号を返す
' keyword : I : 検索ワード
' input_array : I : 文字列の配列
' is_use_regexp : I : 正規表現の使用有無
' Ret : ヒットした行番号
'-------------------------------------------------------------
Public Function FindRowByKeywordFromArray(ByVal keyword As String, ByRef input_array() As String, ByVal is_use_regexp As Boolean) As Long
    If keyword = "" Then
        FindRowByKeywordFromArray = -1
        Exit Function
    End If

    Dim row As Long
    Dim isMatch As Boolean
    Dim line As String
    Dim regex As Object
    Set regex = Nothing
    
    If is_use_regexp = True Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = keyword
    End If
   
    For row = LBound(input_array) To UBound(input_array)
        line = input_array(row)
        
        If is_use_regexp = True Then
            isMatch = regex.Test(line)
        ElseIf InStr(1, line, keyword) > 0 Then
            isMatch = True
        End If
    
        If isMatch = True Then
            FindRowByKeywordFromArray = row
            Exit Function
        End If
    Next row
    
    FindRowByKeywordFromArray = -1
End Function

'-------------------------------------------------------------
'ワークシートの指定列の全行を指定ワードで検索し、ヒットした行番号リストを返す
' ws : I : ワークシート
' find_clm : I : 指定列名(Ex."A")
' find_start_row : I : 検索開始行(1始まり)
' find_end_row : I : 検索終了行(0の場合は全行とする)
' keyword : I : 検索ワード
' Ret : ヒットした行番号リスト
'-------------------------------------------------------------
Public Function FindRowListByKeywordFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal find_clm As String, _
  ByVal find_start_row As Long, _
  ByVal find_end_row As Long, _
  ByVal keyword As String _
) As Long()
    Dim rng As Range
    Dim cell As Range
    Dim cnt As Long
    Dim found_rows() As Long
    
    If find_start_row > find_end_row Or find_end_row < 0 Then
        find_end_row = 0
    End If
    
    If find_end_row = 0 Then
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlUp).row)
    Else
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlUp).row)
    End If

    cnt = 0
    
    For Each cell In rng
        If cell.value = keyword Then
            ReDim Preserve found_rows(cnt)
            found_rows(cnt) = cell.row
            cnt = cnt + 1
        End If
    Next cell
    
    FindRowListByKeywordFromWorksheet = found_rows
End Function

'-------------------------------------------------------------
'ワークシートの指定列の全行を指定ワードで検索し、ヒットした行番号を返す
' ws : I : ワークシート
' find_clm : I : 指定列名(Ex."A")
' find_start_row : I : 検索開始行(1始まり)
' keyword : I : 検索ワード
' find_end_row : I : 検索終了行(任意。0の場合は全行とする)
' is_ignore_case : I : 大文字小文字を区別する(任意。True=区別する(デフォルト), False=区別しない)
' Ret : ヒットした行番号
'-------------------------------------------------------------
Public Function FindRowByKeywordFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal find_clm As String, _
  ByVal find_start_row As Long, _
  ByVal keyword As String, _
  Optional ByVal find_end_row As Long = 0, _
  Optional ByVal is_ignore_case As Boolean = True _
) As Long
    Dim rng As Range
    Dim cell As Range
    Dim found_row As Long
    
    If find_start_row > find_end_row Or find_end_row < 0 Then
        find_end_row = 0
    End If
    
    If find_end_row = 0 Then
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlUp).row)
        'Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlDown).row)
    Else
        'Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlUp).row)
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlDown).row)
    End If

    found_row = 0
    For Each cell In rng
        If is_ignore_case = True Then
            If cell.value = keyword Then
                found_row = cell.row
                Exit For
            End If
        Else
            '大文字小文字を区別しない
            If UCase(cell.value) = UCase(keyword) Then
                found_row = cell.row
                Exit For
            End If
        End If
    Next cell
    
    FindRowByKeywordFromWorksheet = found_row
End Function

'-------------------------------------------------------------
'シートの内容を2次元配列に格納する
' sheet_name : I : シート名
' Ret : シートの内容
'-------------------------------------------------------------
Public Function GetSheetContentsByStringArray(ByVal sheet_name As String) As String()
    Dim ws As Worksheet
    Dim arr() As String
    Dim row_cnt As Long, clm_cnt As Long
    Dim r As Long, c As Long
    
    Set ws = ActiveWorkbook.Worksheets(sheet_name)

    row_cnt = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    clm_cnt = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ReDim arr(1 To row_cnt, 1 To clm_cnt)
    For r = 1 To row_cnt
        For c = 1 To clm_cnt
            arr(r, c) = CStr(ws.Cells(r, c).value)
        Next c
    Next r

    GetSheetContentsByStringArray = arr
End Function

'-------------------------------------------------------------
'拡張子を変更する
' path : I : ファイルパス(絶対パス)
' ext : I : 変更後の拡張子(Ex. ".new")
' Ret : 変更後のファイルパス(絶対パス)
'       pathのファイルが存在しない場合はpathを返す
'-------------------------------------------------------------
Public Function ChangeFileExt(ByVal path As String, ByVal Ext As String) As String
    If IsExistsFile(path) = False Then
        'Err.Raise 53, , "[ChangeFileExt] 指定されたファイルが存在しません (path=" & path & ")"
        ChangeFileExt = path
        Exit Function
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ChangeFileExt] パスが長すぎます (path=" & path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim old_ext As String: old_ext = fso.GetExtensionName(path)
    Dim file_name As String: file_name = fso.GetBaseName(path)
    Dim new_path As String
    
    '新しい拡張子に変更
    file_name = file_name & Ext
    new_path = fso.GetParentFolderName(path) & SEP & file_name
    
    'ファイル名を変更
    fso.MoveFile path, new_path
    Set fso = Nothing
    
    ChangeFileExt = new_path
End Function

'-------------------------------------------------------------
'ブックを開いてシートを取得する
' book_path : I : Excelファイルパス(絶対パス)
' sheet_name : I : シート名
' readonly : I : True/False (True=読取専用で開く, False=読取専用で開かない)
' visible : I : True/False (True=表示, False=非表示)
' Ret : シートオブジェクト
'-------------------------------------------------------------
Public Function GetSheet( _
    ByVal book_path As String, _
    ByVal sheet_name As String, _
    ByVal is_readonly As Boolean, _
    ByVal is_visible As Boolean _
) As Worksheet

    If IsMaxOverPath(book_path) = True Then
        Err.Raise 53, , "[GetSheet] パスが長すぎます (book_path=" & book_path & ")"
    End If

    Dim wb As Workbook
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    
    If IsOpenWorkbook(book_path) = True Then
        '既に開いている
        Set wb = Workbooks(book_path)
    Else
        Set wb = Workbooks.Open(filename:=book_path, UpdateLinks:=False, ReadOnly:=is_readonly)
    End If
    
    wb.Activate
    ActiveWindow.Visible = is_visible
    
    If Common.IsExistSheet(wb, sheet_name) = False Then
        Err.Raise 53, , "[GetSheet] 指定されたシートが存在しません (book_path=" & book_path & ", sheet_name=" & sheet_name & ")"
    End If
    
    Set GetSheet = wb.Worksheets(sheet_name)

End Function

'-------------------------------------------------------------
'ブックを保存して閉じる
' name : I : ブック名(Excelファイル名)
' is_textcmp : I : True=大文字・小文字を区別しない, False=大文字・小文字を区別する(デフォルト)
'-------------------------------------------------------------
Public Sub SaveAndCloseBook(ByVal Name As String, Optional ByVal is_textcmp As Boolean = False)
    Dim wb As Workbook
    For Each wb In Workbooks
        If is_textcmp = False Then
            If InStr(wb.Name, Name) > 0 Then
                wb.Save
                wb.Close
            End If
        Else
            If InStr(1, wb.Name, Name, vbTextCompare) > 0 Then
                wb.Save
                wb.Close
            End If
        End If
    Next
End Sub

'-------------------------------------------------------------
'ブックを閉じる
' name : I : ブック名(Excelファイル名)
' is_textcmp : I : True=大文字・小文字を区別しない, False=大文字・小文字を区別する(デフォルト)
'-------------------------------------------------------------
Public Sub CloseBook(ByVal Name As String, Optional ByVal is_textcmp As Boolean = False)
    Dim wb As Workbook
    For Each wb In Workbooks
        If is_textcmp = False Then
            If InStr(wb.Name, Name) > 0 Then
                wb.Close SaveChanges:=False
            End If
        Else
            If InStr(1, wb.Name, Name, vbTextCompare) > 0 Then
                wb.Close SaveChanges:=False
            End If
        End If
    Next
End Sub

'-------------------------------------------------------------
'ファイルを削除する
' path : IN : ファイルパス(絶対パス)
'-------------------------------------------------------------
Public Sub DeleteFile(ByVal path As String)
    If IsExistsFile(path) = False Then
        Exit Sub
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[DeleteFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const DELETE_READONLY = True
    fso.DeleteFile path, DELETE_READONLY
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'ファイル名をユニークな名称にリネームしてコピーする
' src_file_path : IN : コピー元ファイルパス(絶対パス)
' dst_dir_path : IN : コピー先フォルダパス(絶対パス)
'                     末尾の\は不要
'                     空の場合はコピー元と同じフォルダとする
' Ret : リネームコピー後のファイルパス
'-------------------------------------------------------------
Public Function CopyUniqueFile(ByVal src_file_path As String, ByVal dst_dir_path As String) As String
    If IsExistsFile(src_file_path) = False Then
        CopyUniqueFile = ""
        Exit Function
    End If

    If IsMaxOverPath(src_file_path) = True Or IsMaxOverPath(dst_dir_path) = True Then
        Err.Raise 53, , "[CopyUniqueFile] パスが長すぎます (src_file_path=" & src_file_path & ", dst_dir_path=" & dst_dir_path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim dst_file_path As String
    
    Dim unique_filename As String: unique_filename = GetFileName(src_file_path) & ".bak_" & GetNowTimeString()
    
    If dst_dir_path = "" Then
        dst_file_path = GetFolderNameFromPath(src_file_path) & SEP & unique_filename
    Else
        dst_file_path = dst_dir_path & SEP & unique_filename
    End If

    FileCopy src_file_path, dst_file_path
    
    CopyUniqueFile = dst_file_path
End Function

'-------------------------------------------------------------
'ファイル名を返す
' path : IN : ファイルパス(絶対パス)
' Ret : ファイル名
'-------------------------------------------------------------
Public Function GetFileName(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFileName] パスが長すぎます (path=" & path & ")"
    End If
    
    If path = "" Then
        Err.Raise 53, , "[GetFileName] パスが空です (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイルパスからフォルダパスを返す
' path : IN : ファイルパス(絶対パス)
' Ret : フォルダパス(絶対パス)
'-------------------------------------------------------------
Public Function GetFolderPath(ByVal file_path As String) As String
    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[GetFolderPath] パスが長すぎます (file_path=" & file_path & ")"
    End If
    
    If file_path = "" Then
        Err.Raise 53, , "[GetFolderPath] パスが空です (file_path=" & file_path & ")"
    End If
    
    Dim pos As Integer
    pos = InStrRev(file_path, "\")
    
    '\より左側の文字列をフォルダパスとして返す
    GetFolderPath = Left(file_path, pos - 1)
End Function

'-------------------------------------------------------------
'指定フォルダ配下を指定ファイル名で検索してファイルパスを返す
' search_path : IN : 検索フォルダパス(絶対パス)
' search_name : IN : 検索ファイル名
' Ret : ファイルパス
'-------------------------------------------------------------
Public Function SearchFile(ByVal search_path As String, ByVal search_name As String) As String
    If IsMaxOverPath(search_path) = True Then
        Err.Raise 53, , "[SearchFile] パスが長すぎます (search_path=" & search_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(search_path)
    
    Dim file As Object
    For Each file In folder.Files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like search_name Then
            '発見
            SearchFile = file.path
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    'サブフォルダも検索する
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        Dim result As String
        result = SearchFile(subfolder.path, search_name)
        If result <> "" Then
            'サブフォルダから結果が返ってきた場合は、その結果を返す
            SearchFile = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '検索対象のファイルが見つからなかった場合
    SearchFile = ""
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'指定フォルダのUTF8を全てSJISにする
' path : IN : フォルダパス(絶対パス)
' ext : IN : 拡張子(Ex."*.vb")
' is_subdir : IN : サブフォルダ含むか (True=含む)
' Ret : ファイルリスト
'-------------------------------------------------------------
Public Sub UTF8toSJIS_AllFile(ByVal path As String, ByVal Ext As String, ByVal is_subdir As Boolean)
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] 指定されたフォルダが存在しません (path=" & path & ")"
    End If
    
    If Ext = "" Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] 拡張子が指定されていません"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = CreateFileList(path, Ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        UTF8toSJIS src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'指定フォルダのSJISを全てUTF8にする
' path : IN : フォルダパス(絶対パス)
' ext : IN : 拡張子(Ex."*.vb")
' is_subdir : IN : サブフォルダ含むか (True=含む)
' Ret : ファイルリスト
'-------------------------------------------------------------
Public Sub SJIStoUTF8_AllFile(ByVal path As String, ByVal Ext As String, ByVal is_subdir As Boolean)
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] 指定されたフォルダが存在しません (path=" & path & ")"
    End If
    
    If Ext = "" Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] 拡張子が指定されていません"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = CreateFileList(path, Ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        SJIStoUTF8 src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'指定されたファイルをSJIS→UTF8(BOMあり)変換する
' path : IN : ファイルパス(絶対パス)
' is_backup : IN : True/False (True=バックアップする)
'                  →末尾に".bak_現在日時"を付与
'-------------------------------------------------------------
Public Sub SJIStoUTF8(ByVal path As String, ByVal is_backup As Boolean)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[SJIStoUTF8] パスが長すぎます (path=" & path & ")"
    End If

    Dim in_str As String
    Dim buf As String
    Dim i As Long
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS形式のテキストファイルを読み込み
    in_str = ""
    Open path For Input As #filenum
        'テキストをすべて取得する
        Do Until EOF(filenum)
            Line Input #filenum, buf
            in_str = in_str & buf & vbCrLf
        Loop
    Close #filenum
        
    'Shift-JIS以外のファイルを読み込んでしまった場合は終了
    For i = 1 To Len(in_str)
        If Asc(Mid(in_str, i, 1)) = -7295 Then Exit Sub
    Next
    
    'バックアップ
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    'UTF-8（BOM付き）でテキストファイルへ出力
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText in_str, 0
        .SaveToFile path, 2
        .Close
    End With
    
End Sub

'-------------------------------------------------------------
'指定されたファイルをUTF8(BOMあり/なし) → SJIS変換する
' path : IN : ファイルパス(絶対パス)
' is_backup : IN : True/False (True=バックアップする)
'                  →末尾に".bak_現在日時"を付与
'-------------------------------------------------------------
Public Sub UTF8toSJIS(ByVal path As String, ByVal is_backup As Boolean)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[UTF8toSJIS] パスが長すぎます (path=" & path & ")"
    End If

    Dim in_str As String
    Dim out_str() As String
    Dim i As Long
    
    'UTF-8もしくはUTF-8（BOM付き）のテキストファイルを読み込み
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8もしくはUTF-8（BOM付き）以外を読み込んでしまった場合は終了
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                Exit Sub
            End If
        End If
    Next
    
    '改行毎にデータを分ける
    out_str = Split(in_str, vbCrLf)
    
    'バックアップ
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS形式でテキストファイルへ出力
    Open path For Output As #filenum
        For i = 0 To UBound(out_str)
            Print #filenum, out_str(i)
        Next
    Close #filenum

End Sub

'-------------------------------------------------------------
'ファイルがSJISかを判定する
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=SJIS)
'-------------------------------------------------------------
Public Function IsSJIS(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsSJIS] 指定されたファイルが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsSJIS] パスが長すぎます (path=" & path & ")"
    End If

    Dim Ado As Object
    Const TYPE_BINARY = 1
    Set Ado = CreateObject("ADODB.Stream")
    Ado.Type = TYPE_BINARY
    Ado.Open

    Ado.LoadFromFile path
    Dim read_data As String: read_data = Ado.Read
    Ado.Close
    Set Ado = Nothing

    Dim i As Long
    Dim first_byte As Byte
    Dim second_byte As Byte
    Dim is_dbcs As Boolean
    
    For i = 1 To LenB(read_data)

        first_byte = AscB(MidB(read_data, i, 1))

        '全角文字列(DBCS)の先頭1バイトであるか
        is_dbcs = False

        If &H81 <= first_byte And first_byte <= &H9F Then
            is_dbcs = True
        ElseIf &HE0 <= first_byte And first_byte <= &HEF Then
            is_dbcs = True
        End If

        If is_dbcs Then
            i = i + 1

            If i > LenB(read_data) Then
                IsSJIS = False
                Exit Function
            End If

            second_byte = AscB(MidB(read_data, i, 1))

            If &H40 <= second_byte And second_byte <= &H7F Then
                'SJIS!
            ElseIf &H80 <= second_byte And second_byte <= &HFC Then
                'SJIS!
            Else
                IsSJIS = False
                Exit Function
            End If
        End If
    Next

    IsSJIS = True
End Function

'-------------------------------------------------------------
'ファイルがUTF8(BOMあり/なし)かを判定する
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=UTF8(BOMあり/なし))
'-------------------------------------------------------------
Public Function IsUTF8(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8] 指定されたファイルが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsUTF8] パスが長すぎます (path=" & path & ")"
    End If

    Dim in_str As String
    Dim out_str() As String
    Dim i As Long
    
    'UTF-8もしくはUTF-8（BOM付き）のテキストファイルを読み込み
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8もしくはUTF-8（BOM付き）以外を読み込んでしまった場合は終了
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                IsUTF8 = False
                Exit Function
            End If
        End If
    Next
    
    IsUTF8 = True
End Function

'-------------------------------------------------------------
'ファイルがUTF8(BOMあり)かを判定する
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=UTF8(BOMあり), False=左記以外)
'-------------------------------------------------------------
Public Function IsUTF8_WithBom(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8_WithBom] 指定されたファイルが存在しません (path" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsUTF8_WithBom] パスが長すぎます (path=" & path & ")"
    End If

    Dim bytedata() As Byte: bytedata = ReadBinary(path, 3)
    Dim length As Integer: length = UBound(bytedata) + 1
    
    If length < 3 Then
        IsUTF8_WithBom = False
        Exit Function
    End If
    
    If bytedata(0) = &HEF And bytedata(1) = &HBB And bytedata(2) = &HBF Then
        IsUTF8_WithBom = True
    Else
        IsUTF8_WithBom = False
    End If
    
End Function

'-------------------------------------------------------------
'ファイルをバイナリとして指定サイズ読み込む
' path : IN : ファイルパス(絶対パス)
' readsize : IN : 読み込むサイズ
' Ret : 読み込んだバイナリ配列
'-------------------------------------------------------------
Public Function ReadBinary(ByVal path As String, ByVal readsize As Integer) As Byte()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ReadBinary] パスが長すぎます (path=" & path & ")"
    End If

    Dim readdata() As Byte
    
    If readsize <= 0 Then
        ReadBinary = readdata()
        Exit Function
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    Open path For Binary Access Read As #filenum
    
    ReDim readdata(readsize - 1)
    
    Get #filenum, , readdata
    
    Close #filenum
    
    ReadBinary = readdata
End Function

'-------------------------------------------------------------
'指定フォルダ配下に指定拡張子のファイルが存在するか
' path : IN : フォルダパス(絶対パス)
' in_ext : IN : 拡張子(Ex. "*.vb")
' Ret : True/False (True=存在する, False=存在しない)
'-------------------------------------------------------------
Public Function IsExistsExtensionFile(ByVal path As String, ByVal in_ext As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsExtensionFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim Ext As String: Ext = Replace(in_ext, "*", "")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    For Each subfolder In folder.SubFolders
        If IsExistsExtensionFile(subfolder.path, Ext) Then
            Set fso = Nothing
            Set folder = Nothing
            
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next subfolder
    
    For Each file In folder.Files
        If Right(file.Name, Len(Ext)) = Ext Then
            Set fso = Nothing
            Set folder = Nothing
        
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next file
    
    Set fso = Nothing
    Set folder = Nothing

    IsExistsExtensionFile = False
End Function

'-------------------------------------------------------------
'ログファイルをオープンする
' logfile_path : IN : ログファイルパス(絶対パス)
'-------------------------------------------------------------
Public Sub OpenLog(ByVal logfile_path As String)
    If is_log_opened = True Then
        'すでにオープンしているので無視
        Exit Sub
    End If

    If IsMaxOverPath(logfile_path) = True Then
        Err.Raise 53, , "[OpenLog] パスが長すぎます (logfile_path=" & logfile_path & ")"
    End If

    logfile_num = FreeFile()
    Open logfile_path For Append As logfile_num
    is_log_opened = True
End Sub

'-------------------------------------------------------------
'ログファイルに書き込む
' contents : IN : 書き込む内容
'-------------------------------------------------------------
Public Sub WriteLog(ByVal Contents As String)
    If is_log_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & Contents
End Sub

'-------------------------------------------------------------
'ログファイルをクローズする
'-------------------------------------------------------------
Public Sub CloseLog()
    If is_log_opened = False Then
        'オープンされていないので無視
        Exit Sub
    End If
    Close logfile_num
    logfile_num = -1
    is_log_opened = False
End Sub

'-------------------------------------------------------------
'配列の空行を削除する
' arr : IN : 文字列配列
' Ret : 空行を削除した配列
'-------------------------------------------------------------
Public Function DeleteEmptyArray(ByRef arr() As String) As String()
    Dim result() As String
    Dim i As Integer
    Dim count As Integer
    Dim wk As String
    
    If IsEmptyArray(arr) = True Then
        DeleteEmptyArray = result
        Exit Function
    End If
    
    count = 0
    For i = LBound(arr) To UBound(arr)
        wk = Replace(Replace(Replace(arr(i), vbCrLf, ""), vbCr, ""), vbLf, "")
        If wk <> "" Then
            ReDim Preserve result(count)
            result(count) = wk
            count = count + 1
        End If
    Next i
    DeleteEmptyArray = result
End Function

'-------------------------------------------------------------
'ファイルリストを作成する
' path : IN : フォルダパス(絶対パス)
' ext : IN : 拡張子(Ex."*.vb")
' is_subdir : IN : サブフォルダ含むか (True=含む)
' Ret : ファイルリスト(絶対パスのリスト)
'-------------------------------------------------------------
Public Function CreateFileList( _
    ByVal path As String, _
    ByVal Ext As String, _
    ByVal is_subdir As Boolean _
) As String()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFileList] パスが長すぎます (path=" & path & ")"
    End If

    Dim list() As String: list = CreateFileListMain(path, LCase(Ext), is_subdir)
    CreateFileList = FilterFileListByExtension(DeleteEmptyArray(list), LCase(Ext))
End Function

Private Function CreateFileListMain( _
    ByVal path As String, _
    ByVal Ext As String, _
    ByVal is_subdir As Boolean _
) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filelist() As String
    Dim cnt As Integer

    Dim file As String, f As Object
    file = Dir(path & "\" & Ext)
    
    If file <> "" Then
        If IsEmptyArray(filelist) = True Then
            cnt = 0
        Else
            cnt = UBound(filelist) + 1
        End If
        
        ReDim Preserve filelist(cnt)
        filelist(cnt) = path & "\" & file
    End If
    
    Do While file <> ""
        file = Dir()
        If file <> "" Then
            cnt = UBound(filelist) + 1
            ReDim Preserve filelist(cnt)
            filelist(cnt) = path & "\" & file
        End If
    Loop
    
    If is_subdir = False Then
        Set fso = Nothing
        CreateFileListMain = filelist
        Exit Function
    End If
    
    Dim filelist_sub() As String
    Dim filelist_merge() As String
    
    For Each f In fso.GetFolder(path).SubFolders
        filelist_sub = CreateFileListMain(f.path, Ext, is_subdir)
        filelist = MergeArray(filelist_sub, filelist)
    Next f
    
    Set fso = Nothing
    CreateFileListMain = filelist
End Function

'-------------------------------------------------------------
'ファイルパスの配列から指定拡張子のファイルのみを新しい配列にコピーして返す。
' path_list : I : ファイルパスの配列
' in_ext : I : 拡張子(Ex. "*.txt")
' Ret : フィルター後のファイルパスの配列
'-------------------------------------------------------------
Function FilterFileListByExtension(ByRef path_list() As String, in_ext As String) As String()
    Dim i As Long
    Dim j As Long: j = 0
    Dim filtered_list() As String
    Dim Ext As String: Ext = Replace(in_ext, "*", "")
    
    If in_ext = "*.*" Then
        FilterFileListByExtension = path_list
        Exit Function
    End If
    
    If IsEmptyArray(path_list) = True Then
        FilterFileListByExtension = path_list
        Exit Function
    End If
      
    For i = 0 To UBound(path_list)
        If LCase(Right(path_list(i), Len(Ext))) = LCase(Ext) Then
            ReDim Preserve filtered_list(j)
            filtered_list(j) = path_list(i)
            j = j + 1
        End If
    Next i
    
    FilterFileListByExtension = filtered_list
End Function

'-------------------------------------------------------------
'2つの配列を結合して返す
' array1 : IN : 配列1
' array2 : IN : 配列2
' Ret : 結合した配列
'-------------------------------------------------------------
Public Function MergeArray(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim merged As Variant
    merged = Split(Join(array1, vbCrLf) & vbCrLf & Join(array2, vbCrLf), vbCrLf)
    MergeArray = merged
End Function

'-------------------------------------------------------------
'2つのテキストファイルを比較して一致しているかを返す
' file1 : IN : ファイル1パス(絶対パス)
' file2 : IN : ファイル2パス(絶対パス)
' Ret : 比較結果 : True/False (True=一致)
'-------------------------------------------------------------
Public Function IsMatchTextFiles(ByVal file1 As String, ByVal file2 As String) As Boolean
    If IsMaxOverPath(file1) = True Or IsMaxOverPath(file2) = True Then
        Err.Raise 53, , "[IsMatchTextFiles] パスが長すぎます (file1=" & file1 & ", file2=" & file2 & ")"
    End If

    Dim filesize1 As Long: filesize1 = FileLen(file1)
    Dim filesize2 As Long: filesize2 = FileLen(file2)
    
    'TODO:バイナリレベルで比較すべき
    
    'まずファイルサイズでチェック
    If filesize1 = 0 And filesize2 = 0 Then
        'どちらも0byteなので一致
        IsMatchTextFiles = True
        Exit Function
    ElseIf filesize1 <> filesize2 Then
        'ファイルサイズが異なるので不一致
        IsMatchTextFiles = False
        Exit Function
    ElseIf filesize1 = 0 Or filesize2 = 0 Then
        'どちらかが0byteなので不一致
        IsMatchTextFiles = False
        Exit Function
    End If

    Dim fso1, fso2 As Object
    Dim ts1, ts2 As Object
    
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    
    Const READ_ONLY = 1
    Set ts1 = fso1.OpenTextFile(file1, READ_ONLY)
    Set ts2 = fso2.OpenTextFile(file2, READ_ONLY)
    
    Dim contents1 As String: contents1 = ts1.ReadAll
    Dim contents2 As String: contents2 = ts2.ReadAll
    
    ts1.Close
    ts2.Close
    Set ts1 = Nothing
    Set ts2 = Nothing
    Set fso1 = Nothing
    Set fso2 = Nothing
    
    IsMatchTextFiles = (contents1 = contents2)
End Function

'-------------------------------------------------------------
'文字列の配列の末尾に文字列を追加する
' ary : IN/OUT : 文字列の配列
' value : IN : 追加する文字列
'-------------------------------------------------------------
Public Sub AppendArray(ByRef ary() As String, ByVal value As String)
    If IsEmptyArray(ary) = True Then
        ReDim Preserve ary(0)
        ary(0) = value
    Else
        Dim cnt As Integer: cnt = UBound(ary) + 1
        ReDim Preserve ary(cnt)
        ary(cnt) = value
    End If
End Sub

Public Sub AppendArrayLong(ByRef ary() As String, ByVal value As String)
    If IsEmptyArrayLong(ary) = True Then
        ReDim Preserve ary(0)
        ary(0) = value
    Else
        Dim cnt As Long: cnt = UBound(ary) + 1
        ReDim Preserve ary(cnt)
        ary(cnt) = value
    End If
End Sub

'-------------------------------------------------------------
'フォルダパスを列挙する。（サブフォルダ含む）
' 注意：pathは戻り値には含まない
' path : IN : フォルダパス（絶対パス）
' Ret : フォルダパスリスト
'-------------------------------------------------------------
Public Function GetFolderPathList(ByVal path As String) As String()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderPathList] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Dim top_dir As Object
    Dim sub_dir As Object
    Dim path_list() As String
    Dim dir_cnt As Long
    Dim i, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set top_dir = fso.GetFolder(path)

    dir_cnt = top_dir.SubFolders.count
    If dir_cnt > 0 Then
        ReDim path_list(dir_cnt - 1)
        i = 0
        For Each sub_dir In top_dir.SubFolders
            path_list(i) = sub_dir.path
            i = i + 1
            
            Dim sub_path_list() As String
            sub_path_list = GetFolderPathList(sub_dir.path)
            
            'サブフォルダ内のパスを配列に追加する
            If sub_path_list(0) <> "" Then
                Dim cnt As Integer: cnt = UBound(path_list) + UBound(sub_path_list) + 1
                ReDim Preserve path_list(cnt)
                For j = LBound(sub_path_list) To UBound(sub_path_list)
                    path_list(i) = sub_path_list(j)
                    i = i + 1
                Next j
            End If
        Next sub_dir
        
        GetFolderPathList = path_list
    Else
        Dim ret_empty(0) As String
        GetFolderPathList = ret_empty
    End If
    
    Set sub_dir = Nothing
    Set top_dir = Nothing
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'フォルダをコピーする(サブフォルダ含む)
' src_path : IN : コピー元フォルダパス(絶対パス)
' dst_path : IN : コピー先フォルダパス(絶対パス)
'-------------------------------------------------------------
Public Sub CopyFolder(ByVal src_path As String, dest_path As String)
    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dest_path) = True Then
        Err.Raise 53, , "[CopyFolder] パスが長すぎます (src_path=" & src_path & ", dest_path=" & dest_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'コピー元のフォルダが存在しない場合、エラーを発生させる
    If Not fso.FolderExists(src_path) Then
        Err.Raise 53, , "[CopyFolder] 指定されたフォルダが存在しません。(src_path=" & src_path & ")"
    End If
    
    'コピー先のフォルダが存在しない場合、作成する
    If Not fso.FolderExists(dest_path) Then
        CreateFolder dest_path
    End If
    
    'コピー元のフォルダ内のファイルをコピーする
    Const OVERWRITE = True
    Dim file As Object
    For Each file In fso.GetFolder(src_path).Files
        fso.CopyFile file.path, fso.BuildPath(dest_path, file.Name), OVERWRITE
    Next
    
    'コピー元のフォルダ内のサブフォルダをコピーする
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(src_path).SubFolders
        CopyFolder subfolder.path, fso.BuildPath(dest_path, subfolder.Name)
    Next
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'Yes/Noメッセージボックスを表示する
' msg : IN : メッセージ
' Ret : True/False (True=Yes, False=No)
'-------------------------------------------------------------
Public Function ShowYesNoMessageBox(ByVal msg As String) As Boolean
    Dim result As Integer: result = MsgBox(msg, vbYesNo, "Confirm")
    
    If result = vbYes Then
        ShowYesNoMessageBox = True
    Else
        ShowYesNoMessageBox = False
    End If
End Function

'-------------------------------------------------------------
'外部アプリケーションを実行し、終了するまで待機する
' exe_path : IN : 外部アプリケーション(exe)の絶対パス
'                 exeに渡すパラメータがある場合も一緒に書くこと
' Ret : プロセスの戻り値
'-------------------------------------------------------------
Public Function RunProcessWait(ByVal exe_path As String) As Long
    If IsMaxOverPath(exe_path) = True Then
        Err.Raise 53, , "[RunProcessWait] パスが長すぎます (exe_path=" & exe_path & ")"
    End If

    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    
    Const NOT_DISP = 0
    Const DISP = 1
    Const WAIT = True
    Const NO_WAIT = False
    
    Dim Process As Object
    Set Process = wsh.Exec(exe_path)
    
    'プロセス完了時に通知を受け取る
    Do While Process.Status = 0
        DoEvents
    Loop
    
    'プロセスの戻り値を取得する
    RunProcessWait = Process.ExitCode
    
    Set Process = Nothing
    Set wsh = Nothing
End Function

'-------------------------------------------------------------
' BATファイルを実行する
' bat_path : IN : BATファイルの絶対パス
'                 BATに渡すパラメータがある場合も一緒に書くこと
' Ret : BATの戻り値(exit /b 0の場合0が戻る)
'-------------------------------------------------------------
Public Function RunBatFile(ByVal bat_path As String) As Long
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim returnValue As Variant
    
    Const NOT_DISP = 0
    Const DISP = 1
    Const WAIT = True
    Const NO_WAIT = False
    
    returnValue = wsh.Run(bat_path, NOT_DISP, WAIT)
    
    RunBatFile = CLng(returnValue)
    
    Set wsh = Nothing
End Function

'-------------------------------------------------------------
'前後のダブルクォーテーションを除去して返す
' 例:"hoge" → hoge
' target : IN : 対象文字列
' Ret : 除去後の文字列
'-------------------------------------------------------------
Public Function RemoveQuotes(ByVal target As String) As String
    '""で囲まれているかをチェック
    If Left(target, 1) = """" And Right(target, 1) = """" Then
        '""を削除して返す
        RemoveQuotes = Mid(target, 2, Len(target) - 2)
    Else
        RemoveQuotes = target
    End If
End Function

'-------------------------------------------------------------
'パス文字列の末尾の\を除去して返す
' path : IN : パス文字列
' Ret : パス文字列
'-------------------------------------------------------------
Public Function RemoveTrailingBackslash(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RemoveTrailingBackslash] パスが長すぎます (path=" & path & ")"
    End If

    If Right(path, 1) = "\" Then
        path = Left(path, Len(path) - 1)
    End If
    RemoveTrailingBackslash = path
End Function

'-------------------------------------------------------------
'ファイルの内容を指定されたシートに出力する
' file_path : IN : ファイルパス (絶対パス)
' sheet_name : IN : シート名
'-------------------------------------------------------------
Public Sub OutputTextFileToSheet(ByVal file_path As String, ByVal sheet_name As String)
    If IsExistsFile(file_path) = False Or sheet_name = "" Then
        Err.Raise 53, , "[OutputTextFileToSheet] 指定されたファイルが存在しません (file_path=" & file_path & ")"
    End If

    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[OutputTextFileToSheet] パスが長すぎます (file_path=" & file_path & ")"
    End If

    'ワーク用にコピーする
    Dim wk As String: wk = CopyUniqueFile(file_path, "")
    
    'ワークファイルをSJISに変換する
    UTF8toSJIS wk, False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'ファイルを開く
    Const FORMAT_ASCII = 0
    Const FORMAT_UNICODE = -1
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    'ファイルの内容をシートに出力
    Dim row As Integer: row = 1
    
    Do While Not fileobj.AtEndOfStream
        ws.Cells(row, 1).value = fileobj.ReadLine
        row = row + 1
    Loop
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    'ワークファイルを削除する
    DeleteFile wk
End Sub

'-------------------------------------------------------------
'SJISでテキストファイルを作成する
' contents : IN : 内容
' path : IN : ファイルパス (絶対パス)
'-------------------------------------------------------------
Public Sub CreateSJISTextFile(ByRef Contents() As String, ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateSJISTextFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim txt As Object
    Dim i As Long
    
    Dim is_overwrite As Boolean: is_overwrite = True
    Dim IS_UNICODE As Boolean: IS_UNICODE = False
    
    Set txt = fso.CreateTextFile(path, is_overwrite, IS_UNICODE)
    
    For i = LBound(Contents) To UBound(Contents)
        txt.WriteLine Contents(i)
    Next i
    
    txt.Close
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'ファイルを作成する
' path : IN : ファイルパス (絶対パス)
'-------------------------------------------------------------
Public Sub CreateFile(ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFile] パスが長すぎます (path=" & path & ")"
    End If
    
    If path = "" Then
        Err.Raise 53, , "[CreateFile] パスが指定されていません (path=" & path & ")"
    End If
    
    If IsExistsFile(path) = True Then
        Exit Sub
    End If
    
    Dim dir_path As String
    dir_path = Common.GetFolderPath(path)
    
    If Common.IsExistsFolder(dir_path) = False Then
        CreateFolder dir_path
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CreateTextFile path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'サブフォルダをまとめて作成する
' path : IN : フォルダパス (絶対パス)
'-------------------------------------------------------------
Public Sub CreateFolder(ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFolder] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folders() As String
    folders = Split(path, Application.PathSeparator)
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(folders) To UBound(folders)
        ary = folders
        ReDim Preserve ary(i)
        If Not fso.FolderExists(Join(ary, Application.PathSeparator)) Then
            Call fso.CreateFolder(Join(ary, Application.PathSeparator))
        End If
    Next
  
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'フォルダを削除する
' path : IN : フォルダパス (絶対パス)
'-------------------------------------------------------------
Public Sub DeleteFolder(ByVal path As String)
    If IsExistsFolder(path) = False Then
        Exit Sub
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[DeleteFolder] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.DeleteFolder path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'フォルダを移動する
' src_path : IN : 移動元フォルダパス (絶対パス)
' dst_path : IN : 移動先フォルダパス (絶対パス)
'-------------------------------------------------------------
Public Sub MoveFolder(ByVal src_path As String, ByVal dst_path As String)
    If IsExistsFolder(src_path) = False Then
        Err.Raise 53, , "[MoveFolder] 移動元フォルダが存在しません (src_path=" & src_path & ")"
        Exit Sub
    End If

    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dst_path) = True Then
        Err.Raise 53, , "[MoveFolder] パスが長すぎます (src_path=" & src_path & ", dst_path=" & dst_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim err_msg As String
    Dim retry As Integer
    For retry = 0 To 3

On Error Resume Next
        fso.MoveFolder src_path, dst_path
    
        err_msg = Err.Description
        Err.Clear
On Error GoTo 0

        If err_msg = "" Then
            Exit For
        End If
        
        WaitSec 1

    Next retry
    
    Set fso = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , "[MoveFolder] エラー! (err_msg=" & err_msg & ")"
    End If
    
End Sub

'-------------------------------------------------------------
'文字列配列の共通文字列を返す
' list : IN : 文字列配列
' Ret : 共通文字列
'       Ex. list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'           Ret = "hoge"
'-------------------------------------------------------------
Function GetCommonString(ByRef list() As String) As String
    Dim common_string As String
    Dim i, j As Long
    Dim flag As Boolean
    
    '最初の文字列を共通文字列の初期値とする
    common_string = list(0)
    
    '各文字列を比較する
    For i = 1 To UBound(list)
        flag = False
        '共通部分を取得する
        For j = 1 To Len(common_string)
            If Mid(common_string, j, 1) <> Mid(list(i), j, 1) Then
                common_string = Left(common_string, j - 1)
                flag = True
                Exit For
            End If
        Next j
    Next i
    
    '結果を出力する
    GetCommonString = common_string
End Function

'-------------------------------------------------------------
'絶対ファイルパスの親フォルダパスを取得する
' path : IN : ファイルパス (絶対パス)
' Ret : 親フォルダパス (絶対パス)
'       Ex. path = "C:\tmp\abc.txt"
'           Ret = "C:\tmp"
'-------------------------------------------------------------
Public Function GetFolderNameFromPath(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderNameFromPath] パスが長すぎます (path=" & path & ")"
    End If

    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = Left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'-------------------------------------------------------------
'相対パスを絶対パスに変換する
' base_path : IN : 基準となるフォルダパス(絶対パス)
' ref_path : IN : ファイルパス（相対パス)
' Ret : 絶対パス
'       Ex. base_path = "C:\tmp\abc"
'           ref_path = "..\cdf\xyz.txt"
'           Ret = "C:\tmp\cdf\xyz.txt"
'-------------------------------------------------------------
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
    If IsMaxOverPath(base_path) = True Or IsMaxOverPath(ref_path) = True Then
        Err.Raise 53, , "[GetAbsolutePathName] パスが長すぎます (base_path=" & base_path & ", ref_path=" & ref_path & ")"
    End If

     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイルの存在チェックを行う
' path : IN : ファイルパス(絶対パス)
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistsFile(ByVal path As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsFile] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        IsExistsFile = True
    Else
        IsExistsFile = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'フォルダの存在チェックを行う
' path : IN : フォルダパス(絶対パス)
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistsFolder(ByVal path As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsFolder] パスが長すぎます (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'ファイル名から拡張子を返す
' filename : IN : ファイル名
' isRaw    : IN : True=取得した拡張子をそのまま返す、False=小文字に変換して返す(デフォルト)
' Ret : ファイル名の拡張子
'        Ex. "abc.txt"の場合、"txt"が返る
'            "."が含まれていない場合は""が返る
'-------------------------------------------------------------
Public Function GetFileExtension(ByVal filename As String, Optional ByVal isRaw As Boolean = False) As String
    Dim dot_pos As Integer
    
    ' "."の位置を取得
    dot_pos = InStrRev(filename, ".")
    
    ' 拡張子を取得
    If dot_pos > 0 Then
        If isRaw = False Then
            GetFileExtension = LCase(Right(filename, Len(filename) - dot_pos))
        Else
            GetFileExtension = Right(filename, Len(filename) - dot_pos)
        End If
    Else
        GetFileExtension = ""
    End If
End Function

'-------------------------------------------------------------
'指定フォルダ配下を指定ファイル名で検索してその内容を返す
' target_folder : IN :検索フォルダパス(絶対パス)
' target_file : IN :検索ファイル名
' Ret : 読み込んだファイルの内容
'       配列の末尾には検索ファイルの絶対パスを格納する
'-------------------------------------------------------------
Public Function SearchAndReadFiles(ByVal target_folder As String, ByVal target_file As String) As String()
    If IsMaxOverPath(target_folder) = True Then
        Err.Raise 53, , "[SearchAndReadFiles] パスが長すぎます (target_folder=" & target_folder & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim fileobj As Object
    For Each fileobj In folder.Files
        If fso.FileExists(fileobj.path) And fso.GetFileName(fileobj.path) Like target_file Then
            '検索対象のファイルを読み込む
            Dim Contents As String: Contents = ReadTextFileBySJIS(fileobj.path)

            'ファイルの内容を配列に格納する
            Dim lines() As String: lines = Split(Contents, vbCrLf)
            
            '末尾にファイルパスを追加する
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = file.path
            SearchAndReadFiles = lines
            Set fileobj = Nothing
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    'サブフォルダも検索する
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        Dim result() As String
        result = SearchAndReadFiles(subfolder.path, target_file, is_sjis)
        If UBound(result) >= 1 Then
            'サブフォルダから結果が返ってきた場合は、その結果を返す
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '検索対象のファイルが見つからなかった場合は、空の配列を返す
    Dim ret_empty(0) As String
    SearchAndReadFiles = ret_empty
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'SJISでテキストファイルを読み込む
'※UTF8のファイルもSJISに変換して読み込む!
' path : IN : ファイルパス (絶対パス)
' Ret : 読み込んだ内容
'-------------------------------------------------------------
Public Function ReadTextFileBySJIS(ByVal path As String) As String
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[ReadTextFileBySJIS] 指定されたファイルが存在しません (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ReadTextFileBySJIS] パスが長すぎます (path=" & path & ")"
    End If

    'ワーク用にコピーする
    Dim wk As String: wk = CopyUniqueFile(path, "")
    
    'ワークファイルをSJISに変換する
    UTF8toSJIS wk, False
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const FORMAT_ASCII = 0
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    Dim Contents As String: Contents = fileobj.ReadAll
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    'ワークファイルを削除する
    DeleteFile wk
    
    ReadTextFileBySJIS = RTrim(Contents)
End Function

'-------------------------------------------------------------
'UTF-8形式のテキストファイルを読み込む
' file_path : IN : ファイルパス (絶対パス)
' Ret : 読み込んだ内容
'-------------------------------------------------------------
Public Function ReadTextFileByUTF8(ByVal file_path) As String
    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[ReadTextFileByUTF8] パスが長すぎます (file_path=" & file_path & ")"
    End If
    
    Dim Contents As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file_path
        Contents = .ReadText
        .Close
    End With
    
    ReadTextFileByUTF8 = Contents
End Function

'-------------------------------------------------------------
'配列が空かをチェックする
' arr : IN : 配列
' Ret : True/False (True=空)
'-------------------------------------------------------------
Public Function IsEmptyArray(arr As Variant) As Boolean
    On Error Resume Next
    Dim i As Integer
    i = UBound(arr)
    If i >= 0 And Err.Number = 0 Then
        'WriteLog "★IsEmptyArray Faile! (" & Err.Description & ")"
        IsEmptyArray = False
    Else
        IsEmptyArray = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsEmptyArrayLong(arr As Variant) As Boolean
    On Error Resume Next
    Dim i As Long
    i = UBound(arr)
    If i >= 0 And Err.Number = 0 Then
        'WriteLog "★IsEmptyArrayLong Faile! (" & Err.Description & ")"
        IsEmptyArrayLong = False
    Else
        IsEmptyArrayLong = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------
'n秒待つ
' sec : I : 待つ時間(秒) ※小数も可
'-------------------------------------------------------------
Public Sub WaitSec(ByVal sec As Double)
    Application.WAIT [Now()] + sec / 86400
End Sub

'-------------------------------------------------------------
'現在日時をミリ秒単位の文字列で返す
' Ret :Ex."20230326123456001"
'-------------------------------------------------------------
Public Function GetNowTimeString() As String
    Dim t As SYSTEMTIME

    Call GetLocalTime(t)
    
    Dim yyyy As String: yyyy = Format(t.wYear, "0000")
    Dim mm As String: mm = Format(t.wMonth, "00")
    Dim dd As String: dd = Format(t.wDay, "00")
    Dim hh As String: hh = Format(t.wHour, "00")
    Dim mn As String: mn = Format(t.wMinute, "00")
    Dim ss As String: ss = Format(t.wSecond, "00")
    Dim fff As String: fff = Format(t.wMilliseconds, "000")
    
    GetNowTimeString = yyyy & mm & dd & hh & mn & ss & fff
End Function

Public Function GetNowTimeString_OLD() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString_OLD = str_date & str_time
End Function

'-------------------------------------------------------------
'シートの存在チェック
' wb : I : ワークブック
' sheet_name : I : シート名
' Ret : True/False (True=存在する)
'-------------------------------------------------------------
Public Function IsExistSheet(ByRef wb As Workbook, ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.Name = sheet_name Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'-------------------------------------------------------------
'シートを削除する
' wb : I : ワークブック
' sheet_name : I : シート名
'-------------------------------------------------------------
Public Sub DeleteSheet(ByRef wb As Workbook, ByVal sheet_name As String)
    If IsExistSheet(wb, sheet_name) = False Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    wb.Sheets(sheet_name).Delete
    Application.DisplayAlerts = True
End Sub

'-------------------------------------------------------------
'シートを追加する
' wb : I : ワークブック
' sheet_name : I : シート名
'-------------------------------------------------------------
Public Sub AddSheet(ByRef wb As Workbook, ByVal sheet_name As String)
    DeleteSheet wb, sheet_name
    wb.Worksheets.Add.Name = sheet_name
End Sub

'-------------------------------------------------------------
'ブックをアクティブにする
' book_name : IN : ブック名(Excelファイル名)
'-------------------------------------------------------------
Public Sub ActiveBook(ByVal book_name As String)
    If IsOpenWorkbook(book_name) = False Then
        Err.Raise 53, , "[ActiveBook] ブックが開かれていません (book_name=" & book_name & ")"
    End If
    
    Dim wb As Workbook
    Set wb = Workbooks(book_name)
    wb.Activate
End Sub

'-------------------------------------------------------------
'指定されたシートの指定セルに値を出力する
' book_name : IN : ワークブック
' sheet_name : IN : シート名
' cell_row : 行
' cell_clm : 列
' contents : IN : 出力する内容
'-------------------------------------------------------------
Public Sub UpdateSheet( _
    ByRef book_name As Workbook, _
    ByVal sheet_name As String, _
    ByVal cell_row As Long, ByVal cell_clm As Long, _
    ByVal Contents As String)
    
    If IsExistSheet(book_name, sheet_name) = False Then
        Err.Raise 53, , "[UpdateSheet] シートが見つかりません (book_name=" & book_name & "), sheet_name=" & sheet_name & ")"
    End If
    
    If cell_row < 0 Or cell_clm < 0 Then
        Err.Raise 53, , "[UpdateSheet] セル位置が不正です (cell_row=" & cell_row & "), cell_clm=" & cell_clm & ")"
    End If
    
    Dim ws As Worksheet
    Set ws = book_name.Sheets(sheet_name)
    
    ws.Cells(cell_row, cell_clm).value = Contents
End Sub






