var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";
var METHOD_LIST_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\expectdata\\methodlist_tagjump.tsv";

// 現在行の文字列を取得
var lineText = GetLineStr(0);

// 末尾の改行文字を削除
lineText = lineText.replace(/\r?\n$/, '');

// ダブルクォーテーションをエスケープ（"を\"に変換）
lineText = lineText.replace(/"/g, '\\"');

// METHOD_LIST_PATHもエスケープ（念のため）
var escapedMethodListPath = METHOD_LIST_PATH.replace(/"/g, '\\"');

// コマンドライン引数として渡す（ダブルクォートで囲む）
var args = '"' + lineText + '"' + ' ' + '"' + escapedMethodListPath + '"';

// EXEを実行
ExecCommand('"' + EXE_PATH + '" ' + args);
