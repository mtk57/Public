VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchShapeForm 
   Caption         =   "�e�L�X�g����"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5460
   OleObjectBlob   =   "SearchShapeForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SearchShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VER = "1.1.0" ' �o�[�W�������X�V

' --- �ϐ���` ---
Private lastSearchTerm As String
' ���ύX�_: Shape��Range�̗������i�[���邽�߁A�ϐ�����foundItems�ɕύX
Private foundItems As Collection
Private currentShapeIndex As Long
' ���ǉ�: ���K�\���I�u�W�F�N�g���i�[����ϐ�
Private regex As Object

' �t�H�[�������������ꂽ�Ƃ��̏���
Private Sub UserForm_Initialize()
    ' �t�H�[���̃L���v�V�����Ƀo�[�W������ǉ�
    Me.Caption = "�}�`/�Z�����̕�����̌���/�u�� (ver " & VER & ")"
    
    ' ���W�I�{�^���̃f�t�H���g���u���݂̃V�[�g�̂݁v�ɐݒ�
    Me.optCurrentSheet.Value = True
    
    ' ���ǉ�: �V�����`�F�b�N�{�b�N�X�̃f�t�H���g�l��ݒ�
    Me.chkSearchInCells.Value = True ' �f�t�H���g�ŃZ���������Ώۂɂ���
    Me.chkUseRegex.Value = False     ' �f�t�H���g�Ő��K�\���͎g�p���Ȃ�
    
    ' ���ǉ�: ���K�\���I�u�W�F�N�g��������
    Set regex = CreateObject("VBScript.RegExp")
End Sub

' --- �w���p�[�֐� ---

' �}�`����e�L�X�g�����S�Ɏ擾���邽�߂̐�p�֐��i�ύX�Ȃ��j
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

' ���ǉ�: �e�L�X�g�����������Ɉ�v���邩�𔻒肷��֐�
Private Function IsMatch(ByVal inputText As String, ByVal searchTerm As String) As Boolean
    If Len(inputText) = 0 Or Len(searchTerm) = 0 Then
        IsMatch = False
        Exit Function
    End If

    If Me.chkUseRegex.Value Then
        With regex
            .Pattern = searchTerm
            .Global = True
            .MultiLine = True
            .IgnoreCase = True ' vbTextCompare����
            IsMatch = .Test(inputText)
        End With
    Else
        IsMatch = (InStr(1, inputText, searchTerm, vbTextCompare) > 0)
    End If
End Function

' ���ǉ�: ���K�\�����l�������e�L�X�g�u���֐�
Private Function DoReplace(ByVal inputText As String, ByVal searchTerm As String, ByVal replaceTerm As String) As String
    If Me.chkUseRegex.Value Then
        With regex
            .Pattern = searchTerm
            .Global = True
            .MultiLine = True
            .IgnoreCase = True ' vbTextCompare����
            DoReplace = .Replace(inputText, replaceTerm)
        End With
    Else
        DoReplace = Replace(inputText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
    End If
End Function

' --- �����v���V�[�W�� ---

' �}�`���ċA�I�Ɍ������邽�߂̃v���V�[�W��
Private Sub SearchShapesRecursive(ByVal shapesToSearch As Object, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next
    Dim shp As Shape
    Dim shapeText As String
    For Each shp In shapesToSearch
        If shp.Type = msoGroup Then
            SearchShapesRecursive shp.GroupItems, searchTerm, results
        Else
            shapeText = GetShapeText(shp)
            ' ���ύX�_: IsMatch�֐��Ō���
            If IsMatch(shapeText, searchTerm) Then
                results.Add shp
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

' ���ǉ�: �Z�����������邽�߂̃v���V�[�W��
Private Sub SearchCells(ByVal sheetToSearch As Worksheet, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next
    Dim cell As Range
    For Each cell In sheetToSearch.UsedRange.Cells
        If Not cell.HasFormula Then
            ' ���ύX�_: IsMatch�֐��Ō���
            If IsMatch(cell.Text, searchTerm) Then
                results.Add cell
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

' ���������s���鋤�ʃv���V�[�W��
Private Sub ExecuteSearch()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' �������ʂ�������
    Set foundItems = New Collection
    currentShapeIndex = 0
    lastSearchTerm = searchTerm

    Dim ws As Worksheet
    If Me.optAllSheets.Value = True Then
        ' ���ׂẴV�[�g������
        For Each ws In ActiveWorkbook.Worksheets
            SearchShapesRecursive ws.Shapes, searchTerm, foundItems
            ' ���ǉ�: �Z��������������ꍇ
            If Me.chkSearchInCells.Value Then
                SearchCells ws, searchTerm, foundItems
            End If
        Next ws
    Else
        ' ���݂̃V�[�g�̂݌���
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundItems
        ' ���ǉ�: �Z��������������ꍇ
        If Me.chkSearchInCells.Value Then
            SearchCells ActiveSheet, searchTerm, foundItems
        End If
    End If
End Sub

' --- �C�x���g�n���h�� ---

' �e�L�X�g�{�b�N�X��Enter�L�[�������ꂽ�Ƃ��̏���
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    Call btnSearch_Click
End Sub

' �u���������v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnSearch_Click()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    If Len(searchTerm) = 0 Then Exit Sub
    
    ' --- �������� ---
    If lastSearchTerm <> searchTerm Or foundItems Is Nothing Then
        Call ExecuteSearch
    End If
    
    ' --- �������ʂ̕\�� ---
    If foundItems.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundItems.Count Then currentShapeIndex = 1
        
        ' ���ύX�_: Shape��Range�̗�����������悤�ɂ���
        Dim targetItem As Object
        Set targetItem = foundItems(currentShapeIndex)
        
        Dim targetSheet As Worksheet
        On Error Resume Next
        If TypeName(targetItem) = "Range" Then
            Set targetSheet = targetItem.Worksheet
        Else ' Shape
            Set targetSheet = targetItem.TopLeftCell.Worksheet
        End If
        targetSheet.Activate
        
        ' --- �X�N���[������ ---
        Application.GoTo Reference:=targetItem, Scroll:=True
        If Err.Number <> 0 Then
            Err.Clear
            targetItem.Select
            With ActiveWindow
                If TypeName(targetItem) = "Range" Then
                    .ScrollRow = targetItem.Row
                    .ScrollColumn = targetItem.Column
                Else ' Shape
                    .ScrollRow = targetItem.TopLeftCell.Row
                    .ScrollColumn = targetItem.TopLeftCell.Column
                End If
            End With
        End If
        On Error GoTo 0
    Else
        Beep
    End If

    AppActivate Me.Caption
    Me.txtSearch.SetFocus
End Sub

' �u�u���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnReplace_Click()
    If MsgBox("���݂̍��ڂ�u�����܂����H", vbYesNo + vbQuestion, "�u���̊m�F") = vbNo Then Exit Sub
    
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Or foundItems Is Nothing Or foundItems.Count = 0 Then
        Beep
        Exit Sub
    End If

    ' ���ύX�_: Shape��Range�̗�����������悤�ɂ���
    Dim targetItem As Object
    Set targetItem = foundItems(currentShapeIndex)

    Dim originalText As String, newText As String
    If TypeName(targetItem) = "Range" Then
        originalText = CStr(targetItem.Value)
        newText = DoReplace(originalText, searchTerm, replaceTerm)
        targetItem.Value = newText
    Else ' Shape
        originalText = GetShapeText(targetItem)
        If IsMatch(originalText, searchTerm) Then
            newText = DoReplace(originalText, searchTerm, replaceTerm)
            targetItem.TextFrame.Characters.Text = newText
        End If
    End If
    
    Call btnSearch_Click
End Sub

' �u���ׂĒu���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnReplaceAll_Click()
    Dim scopeText As String
    If Me.optAllSheets.Value = True Then scopeText = "���ׂẴV�[�g" Else scopeText = "���݂̃V�[�g"

    Dim confirmMsg As String
    confirmMsg = "�����͈́F�u" & scopeText & "�v" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "�u" & Me.txtSearch.Text & "�v�����ׂāu" & Me.txtReplace.Text & "�v�ɒu�����܂��B" & vbCrLf
    confirmMsg = confirmMsg & "���̑���͌��ɖ߂��܂���B��낵���ł����H"
    If MsgBox(confirmMsg, vbYesNo + vbExclamation, "���ׂĂ̒u���̊m�F") = vbNo Then Exit Sub

    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text
    If Len(searchTerm) = 0 Then Exit Sub

    If foundItems Is Nothing Or lastSearchTerm <> searchTerm Then Call ExecuteSearch

    If foundItems.Count > 0 Then
        Dim item As Object
        Dim replacedCount As Long
        replacedCount = 0
        On Error Resume Next
        For Each item In foundItems
            Dim originalText As String, newText As String
            If TypeName(item) = "Range" Then
                originalText = CStr(item.Value)
                If IsMatch(originalText, searchTerm) Then
                    item.Value = DoReplace(originalText, searchTerm, replaceTerm)
                    replacedCount = replacedCount + 1
                End If
            Else ' Shape
                originalText = GetShapeText(item)
                If IsMatch(originalText, searchTerm) Then
                    item.TextFrame.Characters.Text = DoReplace(originalText, searchTerm, replaceTerm)
                    replacedCount = replacedCount + 1
                End If
            End If
        Next item
        On Error GoTo 0

        MsgBox replacedCount & "�̍��ڂ�u�����܂����B", vbInformation
        
        Set foundItems = Nothing
        lastSearchTerm = ""
        currentShapeIndex = 0
    Else
        MsgBox "�u���Ώۂ̍��ڂ�������܂���ł����B", vbExclamation
    End If
End Sub

' ���ǉ�: �V�����`�F�b�N�{�b�N�X�̃N���b�N�C�x���g�Ō������ʂ����Z�b�g
Private Sub chkSearchInCells_Click()
    Set foundItems = Nothing
End Sub

Private Sub chkUseRegex_Click()
    Set foundItems = Nothing
End Sub

' �����͈͂̃I�v�V�������ύX���ꂽ��A�������ʂ����Z�b�g����
Private Sub optAllSheets_Click()
    Set foundItems = Nothing
End Sub

Private Sub optCurrentSheet_Click()
    Set foundItems = Nothing
End Sub

' �u�I���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnClose_Click()
    Unload Me
End Sub

' �t�H�[����������Ƃ��̏���
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundItems = Nothing
    ' ���ǉ�: �I�u�W�F�N�g�����
    Set regex = Nothing
End Sub
