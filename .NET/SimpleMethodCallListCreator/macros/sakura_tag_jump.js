var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

// exe�t�@�C���̑��݃`�F�b�N
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

if (!fso.FileExists(EXE_PATH)) {
    shell.Popup("���s�t�@�C����������܂���:\n" + EXE_PATH, 0, "�G���[", 16);
} else {
    // �ȍ~�̏���...

    // ���ݍs�̕�������擾
    var lineText = GetLineStr(0);

    // �����̉��s�������폜
    lineText = lineText.replace(/\r?\n$/, '');

    // �_�u���N�H�[�e�[�V�������G�X�P�[�v�i"��\"�ɕϊ��j
    lineText = lineText.replace(/"/g, '\\"');

    // �R�}���h���C�������Ƃ��ēn���i�_�u���N�H�[�g�ň͂ށj
    var args = '"' + lineText + '"'

    // EXE�����s
    ExecCommand('"' + EXE_PATH + '" ' + args);
}