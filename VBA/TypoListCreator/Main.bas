Attribute VB_Name = "Main"
' ---�y�`�F�b�N���X�g�쐬�x���c�[�� ��������z---

' ���C�������F�V�[�g����P��𒊏o���A���X�g���쐬����
Sub CreateChecklistHelper()
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim wordDic As Object
    Dim targetCell As Range
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    ' --- �����ݒ� ---
    Const LIST_SHEET_NAME As String = "�P�ꃊ�X�g"
    
    Set wsTarget = ActiveSheet
    
    ' �P�ꃊ�X�g�V�[�g�̑��݊m�F
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET_NAME)
    On Error GoTo 0
    If wsList Is Nothing Then
        MsgBox "�V�[�g�u" & LIST_SHEET_NAME & "�v��������܂���B", vbCritical
        Exit Sub
    End If
    
    ' �`�F�b�N�ΏۃV�[�g���s�K�؂ȏꍇ�͒��f
    If wsTarget.Name = "�`�F�b�N���X�g" Or wsTarget.Name = LIST_SHEET_NAME Then
        MsgBox "���̃V�[�g�͒P�ꒊ�o�̑ΏۊO�ł��B", vbExclamation
        Exit Sub
    End If
    
    ' --- �P��̒��o ---
    ' Dictionary�I�u�W�F�N�g�ŒP��ƕp�x���Ǘ�
    Set wordDic = CreateObject("Scripting.Dictionary")
    ' ���K�\���I�u�W�F�N�g�ŃJ�^�J�i��𒊏o
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .Pattern = "[�A-���[]{3,}" '3�����ȏ�̃J�^�J�i��𒊏o�ΏۂƂ���
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
    
    ' --- ���X�g�̏o�� ---
    wsList.Cells.Clear
    wsList.Range("A1").Value = "���o���ꂽ�P��"
    wsList.Range("B1").Value = "�o����"
    wsList.Range("D1").Value = "�y����G���A�z"
    wsList.Range("D2").Value = "1. ���̃��X�g���琄������\�L����I��ŃR�s�["
    wsList.Range("D3").Value = "2. ����(E3�Z��)�ɓ\��t��"
    wsList.Range("E3").Interior.Color = RGB(200, 255, 200) ' E3�Z�����킩��₷���ΐF��
    wsList.Range("D4").Value = "3. ���̃��X�g����\�L������ȏ�I��"
    wsList.Range("D5").Value = "4. �E�̃{�^�����N���b�N �� ���[����'�`�F�b�N���X�g'�ɒǉ�����܂�"
    
    ' ���o���ʂ������o��
    i = 2
    Dim key As Variant
    For Each key In wordDic.Keys
        wsList.Cells(i, 1).Value = key
        wsList.Cells(i, 2).Value = wordDic(key)
        i = i + 1
    Next key
    
    ' �����ڂ𐮂���
    wsList.Columns("A:E").AutoFit
    wsList.Activate

    ' �{�^����ݒu
    Dim btn As Button
    On Error Resume Next
    wsList.Buttons("AddRuleButton").Delete
    On Error GoTo 0
    Set btn = wsList.Buttons.Add(wsList.Range("E4").Left, wsList.Range("E4").Top, wsList.Range("E4:F4").Width, wsList.Range("E4").Height)
    With btn
        .OnAction = "AddRuleToChecklist"
        .Caption = "�I�������\�L�������[���ɒǉ�"
        .Name = "AddRuleButton"
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "�P��̒��o���������܂����B�u" & LIST_SHEET_NAME & "�v�V�[�g���m�F���Ă��������B", vbInformation
End Sub

' �{�^������Ăяo�����}�N���F�I�����ꂽ�P�ꂩ�烋�[�����쐬����
Sub AddRuleToChecklist()
    Dim wsList As Worksheet
    Dim wsChecklist As Worksheet
    Dim recommendedWord As String
    Dim problemWord As Range
    Dim nextRow As Long
    
    ' --- �����ݒ� ---
    Const LIST_SHEET_NAME As String = "�P�ꃊ�X�g"
    Const CHECKLIST_SHEET_NAME As String = "�`�F�b�N���X�g"
    
    Set wsList = ThisWorkbook.Worksheets(LIST_SHEET_NAME)
    Set wsChecklist = ThisWorkbook.Worksheets(CHECKLIST_SHEET_NAME)
    
    ' �����\�L�����͂���Ă��邩�`�F�b�N
    recommendedWord = wsList.Range("E3").Value
    If recommendedWord = "" Then
        MsgBox "��������\�L��E3�Z���ɓ��͂���Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �I��͈͂�A�񂩃`�F�b�N
    If Selection.Column <> 1 Then
        MsgBox "A��̒P�ꃊ�X�g����A�\�L����I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' --- ���[���ǉ����� ---
    ' �`�F�b�N���X�g�V�[�g�̒ǋL����s���擾
    nextRow = wsChecklist.Cells(wsChecklist.Rows.Count, "A").End(xlUp).Row + 1
    
    For Each problemWord In Selection
        If problemWord.Value <> recommendedWord Then ' �����\�L�ƈقȂ�ꍇ�̂ݒǉ�
            wsChecklist.Cells(nextRow, "A").Value = problemWord.Value
            wsChecklist.Cells(nextRow, "B").Value = recommendedWord
            nextRow = nextRow + 1
        End If
    Next problemWord
    
    wsChecklist.Activate
    MsgBox "�`�F�b�N���X�g�Ƀ��[����ǉ����܂����B", vbInformation
End Sub
' ---�y�`�F�b�N���X�g�쐬�x���c�[�� �����܂Łz---
