var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";
var METHOD_LIST_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\expectdata\\methodlist_tagjump.tsv";

// 現在行の文字列を取得
var lineText = GetLineStr(0);

// 末尾の改行文字を削除
lineText = lineText.replace(/\r?\n$/, '');

// コマンドライン引数として渡す（ダブルクォートで囲む）
var args = '"' + lineText + '"' + ' ' + '"' + METHOD_LIST_PATH + '"';

// EXEを実行
ExecCommand(exePath + " " + args);