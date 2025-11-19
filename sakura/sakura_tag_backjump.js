var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// ユーザーの一時フォルダを取得
var tempFolder = shell.ExpandEnvironmentStrings("%TEMP%");
var workFile = tempFolder + "\\SimpleMethodCallListCreator.tmp";

// ファイルの存在チェック
if (fso.FileExists(workFile)) {
    try {
        // ① ファイル全体を読み込む
        var file = fso.OpenTextFile(workFile, 1); // 1=ForReading
        var fileContent = "";
        // ファイルが空でない場合のみ読み込む
        if (!file.AtEndOfStream) {
            fileContent = file.ReadAll();
        }
        file.Close();

        // ② 行ごとに配列に分割（改行コードでsplit）
        // ※Windowsの改行(\r\n)やUnix(\n)に対応
        var lines = fileContent.split(/\r?\n/);

        // 空行を除去（末尾の改行などで空要素ができるため）
        var validLines = [];
        for (var i = 0; i < lines.length; i++) {
            if (lines[i] && lines[i].length > 0) {
                validLines.push(lines[i]);
            }
        }

        // ③ スタックにデータがあるか確認
        if (validLines.length > 0) {
            // 配列の末尾（最新の情報）を取り出す（Pop）
            var lastEntry = validLines.pop();
            
            // タブ区切りで分割
            var parts = lastEntry.split("\t");

            if (parts.length == 2) {
                var filePath = parts[0];
                var lineNo = parseInt(parts[1], 10);

                if (!isNaN(lineNo) && lineNo > 0) {
                    // ジャンプ実行
                    Editor.FileOpen(filePath);
                    Editor.Jump(lineNo, 1);
                }
            }

            // ④ 履歴情報の更新（使用した最新行を消して書き直す）
            if (validLines.length > 0) {
                // 残りのデータをファイルに書き戻す
                var newContent = validLines.join("\r\n") + "\r\n"; // 末尾に改行をつける
                var newFile = fso.CreateTextFile(workFile, true); // 上書き作成
                newFile.Write(newContent);
                newFile.Close();
            } else {
                // 履歴がなくなった場合はファイルを削除する（クリーンアップ）
                fso.DeleteFile(workFile);
            }
        }
        
    } catch(e) {
        shell.Popup("エラーが発生しました:\n" + e.description, 0, "エラー", 16);
    }
}