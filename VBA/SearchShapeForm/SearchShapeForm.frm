VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchShapeForm 
   Caption         =   "�}�`���e�L�X�g����"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SearchShapeForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SearchShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- SearchShapeForm �̃R�[�h�i�O���[�v���Ή��Łj ---

' �����L�[���[�h���ύX���ꂽ���𔻒f���邽�߂̕ϐ�
Private lastSearchTerm As String
' ���������}�`���i�[����R���N�V����
Private foundShapes As Collection
' ���ɕ\������}�`�̃C���f�b�N�X
Private currentShapeIndex As Long


' �}�`����e�L�X�g�����S�Ɏ擾���邽�߂̐�p�֐�
Private Function GetShapeText(ByVal targetShape As Shape) As String
    On Error Resume Next ' ���̊֐����Ŕ��������G���[�͂��ׂĖ�������
    GetShapeText = "" ' �����l��ݒ�
    
    ' �e�L�X�g�t���[��������A���A�e�L�X�g���܂܂�Ă���ꍇ�ɂ��̓��e��Ԃ�
    If targetShape.HasTextFrame Then
        If targetShape.TextFrame2.HasText Then
            GetShapeText = targetShape.TextFrame.Characters.Text
        End If
    End If
    
    On Error GoTo 0 ' �G���[���������ɖ߂�
End Function


'--- ���������炪�V�K�ǉ������� ---
' �}�`���ċA�I�Ɍ������邽�߂̃v���V�[�W��
Private Sub SearchShapesRecursive(ByVal shapesToSearch As Object, ByVal searchTerm As String, ByRef results As Collection)
    On Error Resume Next ' GroupItems�ȂǂŃG���[���o�邱�Ƃ����邽��

    Dim shp As Shape
    Dim shapeText As String

    For Each shp In shapesToSearch
        ' �}�`���O���[�v�̏ꍇ�A���̒��̃A�C�e���ɑ΂��čēx���̃v���V�[�W�����Ăяo���i�ċA�����j
        If shp.Type = msoGroup Then
            SearchShapesRecursive shp.GroupItems, searchTerm, results
        Else
            ' �O���[�v�łȂ��ꍇ�A�ʏ�ʂ�e�L�X�g������
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
'--- �������܂ł��V�K�ǉ������� ---


' �e�L�X�g�{�b�N�X�ŃL�[�������ꂽ�Ƃ��̏���
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    ' Enter�L�[�������ꂽ�ꍇ�̂ݎ��s
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    ' Enter�L�[�̓��͂𖳌����i�u�|���v�Ƃ�������h���j
    KeyCode = 0
    
    Dim searchTerm As String
    searchTerm = Me.txtSearch.Text
    
    ' �����L�[���[�h����Ȃ牽�����Ȃ�
    If Len(searchTerm) = 0 Then Exit Sub
    
    ' --- �������� ---
    ' �����L�[���[�h���ς�������A���߂Ă̌����̏ꍇ
    If lastSearchTerm <> searchTerm Or foundShapes Is Nothing Then
        lastSearchTerm = searchTerm
        Set foundShapes = New Collection
        currentShapeIndex = 0
        
        '--- ���������W�b�N���ċA�Ăяo���ɕύX�� ---
        SearchShapesRecursive ActiveSheet.Shapes, searchTerm, foundShapes
        '--- ���������W�b�N���ċA�Ăяo���ɕύX�� ---
    End If
    
    ' --- �������ʂ̕\�� ---
    If foundShapes.Count > 0 Then
        currentShapeIndex = currentShapeIndex + 1
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        
        targetShape.Parent.Activate
        
        On Error Resume Next
        Application.Goto Reference:=targetShape, Scroll:=True
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
    
End Sub


' �u�I���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnClose_Click()
    Unload Me
End Sub

' �t�H�[����������Ƃ��̏���
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set foundShapes = Nothing
End Sub
