Attribute VB_Name = "Module1"
' �V���[�g�J�b�g�L�[���炱�̃}�N�����Ăяo��
Public Sub ShowSearchShapeForm()
    ' �t�H�[�������ɊJ����Ă��邩�`�F�b�N
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "SearchShapeForm" Then
            ' ���ɊJ���Ă���Ή������Ȃ��ŏI��
            Exit Sub
        End If
    Next frm

    ' �t�H�[�������[�h���X�ŕ\��
    SearchShapeForm.Show vbModeless
End Sub
