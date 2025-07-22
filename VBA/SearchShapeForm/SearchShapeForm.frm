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
Const VER = "1.0.1"

' �����L�[���[�h���ύX���ꂽ���𔻒f���邽�߂̕ϐ�
Private lastSearchTerm As String
' ���������}�`���i�[����R���N�V����
Private foundShapes As Collection
' ���ɕ\������}�`�̃C���f�b�N�X
Private currentShapeIndex As Long

' �t�H�[�������������ꂽ�Ƃ��̏���
Private Sub UserForm_Initialize()
    ' �t�H�[���̃L���v�V�����Ƀo�[�W������ǉ�
    Me.Caption = "�}�`���̕�����̌���/�u�� (ver " & VER & ")"
    
    ' ���W�I�{�^���̃f�t�H���g���u���݂̃V�[�g�̂݁v�ɐݒ�
    Me.optCurrentSheet.Value = True
End Sub


' �}�`����e�L�X�g�����S�Ɏ擾���邽�߂̐�p�֐�
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
            If Len(shapeText) > 0 Then
                If InStr(1, shapeText, searchTerm, vbTextCompare) > 0 Then
                    results.Add shp
                End If
            End If
        End If
    Next shp
    
    On Error GoTo 0
End Sub


' ���������s���鋤�ʃv���V�[�W��
Private Sub ExecuteSearch()
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' �������ʂ�������
    Set foundShapes = New Collection
    currentShapeIndex = 0
    lastSearchTerm = searchTerm

    If Me.optAllSheets.Value = True Then
        ' ���ׂẴV�[�g������
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            SearchShapesRecursive ws.Shapes, searchTerm, foundShapes
        Next ws
    Else
        ' ���݂̃V�[�g�̂݌���
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
    End If
End Sub


' �e�L�X�g�{�b�N�X�ŃL�[�������ꂽ�Ƃ��̏���
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
    ' �����L�[���[�h���ς�������A���߂Ă̌���(foundShapes����)�̏ꍇ�ɋ��ʌ����������Ăяo��
    If lastSearchTerm <> searchTerm Or foundShapes Is Nothing Then
        Call ExecuteSearch
    End If
    
    ' --- �������ʂ̕\�� ---
    If foundShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        
        ' �}�`���ǂ̃V�[�g��ɂ��邩�𒼐ڒ��ׂāA���̃V�[�g���A�N�e�B�u�ɂ���
        On Error Resume Next
        targetShape.TopLeftCell.Worksheet.Activate
        On Error GoTo 0
        
        ' --- �X�N���[������ ---
        On Error Resume Next
        Application.GoTo Reference:=targetShape, Scroll:=True
        If Err.Number <> 0 Then
            Err.Clear
            targetShape.Select
            With ActiveWindow
                .ScrollRow = targetShape.TopLeftCell.Row
                .ScrollColumn = targetShape.TopLeftCell.Column
            End With
        End If
        On Error GoTo 0
        
    Else
        Beep
    End If

    ' ���ύX�_: AppActivate�Ńt�H�[���������I�ɃA�N�e�B�u�ɂ��A�e�L�X�g�{�b�N�X�Ƀt�H�[�J�X��߂�
    On Error Resume Next ' AppActivate���H�Ɏ��s����ꍇ�����邽�߃G���[�𖳎�����
    AppActivate Me.Caption
    On Error GoTo 0
    Me.txtSearch.SetFocus
End Sub


' �u�u���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnReplace_Click()
    ' �m�F�_�C�A���O�̕\��
    If MsgBox("���݂̐}�`�̃e�L�X�g��u�����܂����H", vbYesNo + vbQuestion, "�u���̊m�F") = vbNo Then
        Exit Sub
    End If
    
    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Or foundShapes Is Nothing Or foundShapes.Count = 0 Then
        Beep
        Exit Sub
    End If

    Dim targetShape As Shape
    Set targetShape = foundShapes(currentShapeIndex)

    Dim originalText As String, newText As String
    originalText = GetShapeText(targetShape)
    If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
        newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
        targetShape.TextFrame.Characters.Text = newText
    End If
    
    Call btnSearch_Click
End Sub


' �u���ׂĒu���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnReplaceAll_Click()
    Dim scopeText As String
    If Me.optAllSheets.Value = True Then
        scopeText = "���ׂẴV�[�g"
    Else
        scopeText = "���݂̃V�[�g"
    End If

    Dim confirmMsg As String
    confirmMsg = "�����͈́F�u" & scopeText & "�v" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "�u" & Me.txtSearch.Text & "�v�����ׂāu" & Me.txtReplace.Text & "�v�ɒu�����܂��B" & vbCrLf
    confirmMsg = confirmMsg & "���̑���͌��ɖ߂��܂���B��낵���ł����H"
    
    If MsgBox(confirmMsg, vbYesNo + vbExclamation, "���ׂĂ̒u���̊m�F") = vbNo Then
        Exit Sub
    End If

    Dim searchTerm As String, replaceTerm As String
    searchTerm = Me.txtSearch.Text
    replaceTerm = Me.txtReplace.Text

    If Len(searchTerm) = 0 Then Exit Sub

    ' --- �������܂����s����Ă��Ȃ��ꍇ�́A�܂����������s ---
    If foundShapes Is Nothing Or lastSearchTerm <> searchTerm Then
        Call ExecuteSearch
    End If

    ' --- �u������ ---
    If foundShapes.Count > 0 Then
        Dim shp As Shape
        Dim replacedCount As Long
        replacedCount = 0

        On Error Resume Next
        For Each shp In foundShapes
            Dim originalText As String, newText As String
            originalText = GetShapeText(shp)
            
            If InStr(1, originalText, searchTerm, vbTextCompare) > 0 Then
                newText = Replace(originalText, searchTerm, replaceTerm, 1, -1, vbTextCompare)
                shp.TextFrame.Characters.Text = newText
                replacedCount = replacedCount + 1
            End If
        Next shp
        On Error GoTo 0

        MsgBox replacedCount & "�̍��ڂ�u�����܂����B", vbInformation
        
        Set foundShapes = Nothing
        lastSearchTerm = ""
        currentShapeIndex = 0

    Else
        MsgBox "�u���Ώۂ̐}�`��������܂���ł����B", vbExclamation
    End If
End Sub


' �����͈͂̃I�v�V�������ύX���ꂽ��A�������ʂ����Z�b�g����
Private Sub optAllSheets_Click()
    Set foundShapes = Nothing
End Sub

Private Sub optCurrentSheet_Click()
    Set foundShapes = Nothing
End Sub


' �u�I���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnClose_Click()
    Unload Me
End Sub


' �t�H�[����������Ƃ��̏���
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundShapes = Nothing
End Sub
