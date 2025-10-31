var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

// 現在行の文字列を取得
var lineText = GetLineStr(0);

// 末尾の改行文字を削除
lineText = lineText.replace(/\r?\n$/, '');

// ダブルクォーテーションをエスケープ（"を\"に変換）
lineText = lineText.replace(/"/g, '\\"');

// コマンドライン引数として渡す（ダブルクォートで囲む）
var args = '"' + lineText + '"'

// EXEを実行
ExecCommand('"' + EXE_PATH + '" ' + args);
