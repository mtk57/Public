var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// ユーザーの一時フォルダを取得
var tempFolder = shell.ExpandEnvironmentStrings("%TEMP%");
var workFile = tempFolder + "\\SimpleMethodCallListCreator.tmp";

//shell.Popup("ワークファイルパス:\n" + workFile, 0, "デバッグ1", 64);

// ①ファイルの存在チェック
if (fso.FileExists(workFile)) {
    //shell.Popup("ファイルが見つかりました", 0, "デバッグ2", 64);
    
    try {
        // ファイルを読み込み
        var file = fso.OpenTextFile(workFile, 1); // 1=読み取り
        var line = file.ReadLine();
        file.Close();
        
        //shell.Popup("読み込んだ内容:\n" + line, 0, "デバッグ3", 64);
        
        // タブ区切りで分割
        var parts = line.split("\t");
        
        //shell.Popup("分割結果:\n要素数=" + parts.length + "\n[0]=" + parts[0] + "\n[1]=" + parts[1], 0, "デバッグ4", 64);
        
        // フォーマットチェック（2要素あるか）
        if (parts.length == 2) {
            var filePath = parts[0];
            var lineNo = parseInt(parts[1], 10);
            
            //shell.Popup("ファイルパス: " + filePath + "\n行番号: " + lineNo + "\n数値チェック: " + !isNaN(lineNo), 0, "デバッグ5", 64);
            
            // 行番号が数値かチェック
            if (!isNaN(lineNo) && lineNo > 0) {
                //shell.Popup("タグジャンプを実行します", 0, "デバッグ6", 64);
                
                // ②ファイルを開いて指定行にジャンプ
                Editor.FileOpen(filePath);
                Editor.Jump(lineNo, 1);
                
                //shell.Popup("タグジャンプ完了", 0, "デバッグ7", 64);
            } else {
                //shell.Popup("行番号が不正です", 0, "デバッグエラー1", 48);
            }
        } else {
            //shell.Popup("フォーマット不正（要素数が2ではない）", 0, "デバッグエラー2", 48);
        }
    } catch(e) {
        //shell.Popup("エラー発生:\n" + e.description, 0, "デバッグエラー3", 16);
    }
} else {
    //shell.Popup("ファイルが見つかりません", 0, "デバッグエラー4", 48);
}