var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

// exeファイルの存在チェック
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

if (!fso.FileExists(EXE_PATH)) {
    shell.Popup("実行ファイルが見つかりません:\n" + EXE_PATH, 0, "エラー", 16);
} else {
    // 以降の処理...

    // 現在行の文字列を取得
    var lineText = GetLineStr(0);

    // 末尾の改行文字を削除
    lineText = lineText.replace(/\r?\n$/, '');

    // ダブルクォーテーションをエスケープ（"を\"に変換）
    lineText = lineText.replace(/"/g, '\\"');

    // コマンドライン引数として渡す（ダブルクォートで囲む）
    var args = '"' + lineText + '"'

    // ①現在のファイルパスと行番号をワークファイルに書き出す
    var currentFilePath = ExpandParameter("$F"); // 現在のファイルパス
    var currentLineNo = ExpandParameter("$y");   // 現在の行番号
    
    // ユーザーの一時フォルダを取得
    var tempFolder = shell.ExpandEnvironmentStrings("%TEMP%");
    var workFile = tempFolder + "\\SimpleMethodCallListCreator.tmp";
    
    // ファイルに書き出し（タブ区切り）
    var file = fso.CreateTextFile(workFile, true);
    file.WriteLine(currentFilePath + "\t" + currentLineNo);
    file.Close();
    
    // EXEを実行
    ExecCommand('"' + EXE_PATH + '" ' + args);
}