Attribute VB_Name = "Main"
' ---【チェックリスト作成支援ツール ここから】---

' メイン処理：シートから単語を抽出し、リストを作成する
Sub CreateChecklistHelper()
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim wordDic As Object
    Dim targetCell As Range
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    ' --- 初期設定 ---
    Const LIST_SHEET_NAME As String = "単語リスト"
    
    Set wsTarget = ActiveSheet
    
    ' 単語リストシートの存在確認
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET_NAME)
    On Error GoTo 0
    If wsList Is Nothing Then
        MsgBox "シート「" & LIST_SHEET_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' チェック対象シートが不適切な場合は中断
    If wsTarget.Name = "チェックリスト" Or wsTarget.Name = LIST_SHEET_NAME Then
        MsgBox "このシートは単語抽出の対象外です。", vbExclamation
        Exit Sub
    End If
    
    ' --- 単語の抽出 ---
    ' Dictionaryオブジェクトで単語と頻度を管理
    Set wordDic = CreateObject("Scripting.Dictionary")
    ' 正規表現オブジェクトでカタカナ語を抽出
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .Pattern = "[ア-ンー]{3,}" '3文字以上のカタカナ語を抽出対象とする
    End With
    
    Application.ScreenUpdating = False
    
    For Each targetCell In wsTarget.UsedRange
        If VarType(targetCell.Value) = vbString And targetCell.Value <> "" Then
            Set matches = regEx.Execute(targetCell.Value)
            For Each match In matches
                If Not wordDic.Exists(match.Value) Then
                    wordDic.Add match.Value, 1
                Else
                    wordDic(match.Value) = wordDic(match.Value) + 1
                End If
            Next match
        End If
    Next targetCell
    
    ' --- リストの出力 ---
    wsList.Cells.Clear
    wsList.Range("A1").Value = "抽出された単語"
    wsList.Range("B1").Value = "出現回数"
    wsList.Range("D1").Value = "【操作エリア】"
    wsList.Range("D2").Value = "1. 下のリストから推奨する表記を一つ選んでコピー"
    wsList.Range("D3").Value = "2. ここ(E3セル)に貼り付け"
    wsList.Range("E3").Interior.Color = RGB(200, 255, 200) ' E3セルをわかりやすく緑色に
    wsList.Range("D4").Value = "3. 下のリストから表記ゆれを一つ以上選択"
    wsList.Range("D5").Value = "4. 右のボタンをクリック → ルールが'チェックリスト'に追加されます"
    
    ' 抽出結果を書き出し
    i = 2
    Dim key As Variant
    For Each key In wordDic.Keys
        wsList.Cells(i, 1).Value = key
        wsList.Cells(i, 2).Value = wordDic(key)
        i = i + 1
    Next key
    
    ' 見た目を整える
    wsList.Columns("A:E").AutoFit
    wsList.Activate

    ' ボタンを設置
    Dim btn As Button
    On Error Resume Next
    wsList.Buttons("AddRuleButton").Delete
    On Error GoTo 0
    Set btn = wsList.Buttons.Add(wsList.Range("E4").Left, wsList.Range("E4").Top, wsList.Range("E4:F4").Width, wsList.Range("E4").Height)
    With btn
        .OnAction = "AddRuleToChecklist"
        .Caption = "選択した表記ゆれをルールに追加"
        .Name = "AddRuleButton"
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "単語の抽出が完了しました。「" & LIST_SHEET_NAME & "」シートを確認してください。", vbInformation
End Sub

' ボタンから呼び出されるマクロ：選択された単語からルールを作成する
Sub AddRuleToChecklist()
    Dim wsList As Worksheet
    Dim wsChecklist As Worksheet
    Dim recommendedWord As String
    Dim problemWord As Range
    Dim nextRow As Long
    
    ' --- 初期設定 ---
    Const LIST_SHEET_NAME As String = "単語リスト"
    Const CHECKLIST_SHEET_NAME As String = "チェックリスト"
    
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET_NAME)
    Set wsChecklist = ThisWorkbook.Worksheets(CHECKLIST_SHEET_NAME)
    
    ' 推奨表記が入力されているかチェック
    recommendedWord = wsList.Range("E3").Value
    If recommendedWord = "" Then
        MsgBox "推奨する表記がE3セルに入力されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 選択範囲がA列かチェック
    If Selection.Column <> 1 Then
        MsgBox "A列の単語リストから、表記ゆれを選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' --- ルール追加処理 ---
    ' チェックリストシートの追記する行を取得
    nextRow = wsChecklist.Cells(wsChecklist.Rows.Count, "A").End(xlUp).Row + 1
    
    For Each problemWord In Selection
        If problemWord.Value <> recommendedWord Then ' 推奨表記と異なる場合のみ追加
            wsChecklist.Cells(nextRow, "A").Value = problemWord.Value
            wsChecklist.Cells(nextRow, "B").Value = recommendedWord
            nextRow = nextRow + 1
        End If
    Next problemWord
    
    wsChecklist.Activate
    MsgBox "チェックリストにルールを追加しました。", vbInformation
End Sub
' ---【チェックリスト作成支援ツール ここまで】---
