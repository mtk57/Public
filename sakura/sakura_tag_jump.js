var EXE_PATH = "C:\\_git\\Public\\.NET\\SimpleMethodCallListCreator\\bin\\Debug\\SimpleMethodCallListCreator.exe";

// exeファイルの存在チェック
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

if (!fso.FileExists(EXE_PATH)) {
    shell.Popup("実行ファイルが見つかりません:\n" + EXE_PATH, 0, "エラー", 16);
} else {
    // 現在行の文字列を取得
    var lineText = GetLineStr(0);

    // 末尾の改行文字を削除
    lineText = lineText.replace(/\r?\n$/, '');

    // ダブルクォーテーションをエスケープ
    lineText = lineText.replace(/"/g, '\\"');

    // コマンドライン引数
    var args = '"' + lineText + '"';

    // ---------------------------------------------------------
    // スタック処理：現在のファイルパスと行番号を追記する
    // ---------------------------------------------------------
    var currentFilePath = ExpandParameter("$F"); // 現在のファイルパス
    var currentLineNo = ExpandParameter("$y");   // 現在の行番号
    
    // ユーザーの一時フォルダを取得
    var tempFolder = shell.ExpandEnvironmentStrings("%TEMP%");
    var workFile = tempFolder + "\\SimpleMethodCallListCreator.tmp";
    
    var file;
    // ファイルが存在するかチェック
    if (fso.FileExists(workFile)) {
        // 存在する場合は「追記モード(8)」で開く
        // 引数: ファイルパス, IOMode(8=Append), Create(false), Format(0=ASCII)
        file = fso.OpenTextFile(workFile, 8, false, 0);
    } else {
        // 存在しない場合は新規作成
        file = fso.CreateTextFile(workFile, true);
    }

    // ファイルに書き出し（タブ区切り）して改行
    file.WriteLine(currentFilePath + "\t" + currentLineNo);
    file.Close();
    
    // EXEを実行
    ExecCommand('"' + EXE_PATH + '" ' + args);
}