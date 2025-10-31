// 現在行の内容を.NET EXEに渡して表示するマクロ

// 現在行の文字列を取得
var lineText = GetLineStr(0);

// 末尾の改行文字を削除
lineText = lineText.replace(/\r?\n$/, '');

// EXEのパスを指定
var exePath = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

// コマンドライン引数として渡す（ダブルクォートで囲む）
var args = '"' + lineText + '"';

// EXEを実行
ExecCommand(exePath + " " + args);