var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

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
