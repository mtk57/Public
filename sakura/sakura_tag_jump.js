var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";
var METHOD_LIST_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\expectdata\\methodlist_tagjump.tsv";

// ���ݍs�̕�������擾
var lineText = GetLineStr(0);

// �����̉��s�������폜
lineText = lineText.replace(/\r?\n$/, '');

// �_�u���N�H�[�e�[�V�������G�X�P�[�v�i"��\"�ɕϊ��j
lineText = lineText.replace(/"/g, '\\"');

// METHOD_LIST_PATH���G�X�P�[�v�i�O�̂��߁j
var escapedMethodListPath = METHOD_LIST_PATH.replace(/"/g, '\\"');

// �R�}���h���C�������Ƃ��ēn���i�_�u���N�H�[�g�ň͂ށj
var args = '"' + lineText + '"' + ' ' + '"' + escapedMethodListPath + '"';

// EXE�����s
ExecCommand('"' + EXE_PATH + '" ' + args);
