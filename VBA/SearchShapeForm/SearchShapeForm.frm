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
' --- SearchShapeForm �̃R�[�h ---

' �����L�[���[�h���ύX���ꂽ���𔻒f���邽�߂̕ϐ�
Private lastSearchTerm As String
' ���������}�`���i�[����R���N�V����
Private foundShapes As Collection
' ���ɕ\������}�`�̃C���f�b�N�X
Private currentShapeIndex As Long

' �u�I���v�{�^���������ꂽ�Ƃ��̏���
Private Sub btnClose_Click()
    Unload Me
End Sub

' �e�L�X�g�{�b�N�X�ŃL�[�������ꂽ�Ƃ��̏����i�G���[�΍��Łj
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
        
        Dim shp As Shape
        Dim hasTextFrameProperty As Boolean
        
        ' �A�N�e�B�u�ȃV�[�g�̂��ׂĂ̐}�`�����[�v
        For Each shp In ActiveSheet.Shapes
            
            '--- ���������炪�C�������� ---
            hasTextFrameProperty = False
            On Error Resume Next ' �ꎞ�I�ɃG���[�𖳎�����
            hasTextFrameProperty = shp.HasTextFrame ' ���̃v���p�e�B�����݂��邩����
            On Error GoTo 0      ' �G���[�����������ɉ�������
            '--- �������܂ł��C�������� ---

            ' HasTextFrame�v���p�e�B������A���e�L�X�g���܂܂�Ă���ꍇ�̂ݏ���
            If hasTextFrameProperty Then
                If shp.TextFrame2.HasText Then
                    ' �e�L�X�g���Ɍ����L�[���[�h���܂܂�Ă��邩�`�F�b�N�i�啶���E����������ʂ��Ȃ��j
                    If InStr(1, shp.TextFrame.Characters.Text, searchTerm, vbTextCompare) > 0 Then
                        foundShapes.Add shp
                    End If
                End If
            End If
            
        Next shp
    End If
    
    ' --- �������ʂ̕\�� ---
    If foundShapes.Count > 0 Then
        ' ���̐}�`�փC���f�b�N�X��i�߂�
        currentShapeIndex = currentShapeIndex + 1
        ' �C���f�b�N�X���}�`�̐��𒴂�����1�ɖ߂�i���[�v����j
        If currentShapeIndex > foundShapes.Count Then
            currentShapeIndex = 1
        End If
        
        ' �Ώۂ̐}�`��I�����āA�����܂ŃX�N���[��
        Dim targetShape As Shape
        Set targetShape = foundShapes(currentShapeIndex)
        targetShape.Parent.Activate
        Application.Goto Reference:=targetShape, Scroll:=True
        
    Else
        ' ������Ȃ������ꍇ
        Beep
    End If
    
End Sub


' �t�H�[����������Ƃ��̏���
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' �ϐ������
    Set foundShapes = Nothing
End Sub
