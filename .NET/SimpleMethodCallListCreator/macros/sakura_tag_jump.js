var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";
var METHOD_LIST_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\expectdata\\methodlist_tagjump.tsv";

// ���ݍs�̕�������擾
var lineText = GetLineStr(0);

// �����̉��s�������폜
lineText = lineText.replace(/\r?\n$/, '');

// �R�}���h���C�������Ƃ��ēn���i�_�u���N�H�[�g�ň͂ށj
var args = '"' + lineText + '"' + ' ' + '"' + METHOD_LIST_PATH + '"';

// EXE�����s
ExecCommand(exePath + " " + args);