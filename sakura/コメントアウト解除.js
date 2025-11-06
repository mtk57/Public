// [マクロ] 選択行のコメントアウト (//) を解除 - 修正版 v3

var nLineFrom = Editor.GetSelectLineFrom();
var nLineTo = Editor.GetSelectLineTo();

if (nLineFrom == 0 || nLineTo == 0) {
    // 処理終了
} else {
    Editor.MoveCursor(nLineFrom, 1, 0); 

    for (var i = nLineFrom; i <= nLineTo; i++) {
        var sLine = Editor.GetLineStr(0); // 現在行の文字列取得
        
        // 先頭の空白をスキップ
        var nPos = 0;
        while (nPos < sLine.length && (sLine.charAt(nPos) == ' ' || sLine.charAt(nPos) == '\t')) {
            nPos++;
        }
        
        // "//" があるかチェック
        if (sLine.substr(nPos, 2) == "//") {
            Editor.MoveCursor(i, nPos + 1, 0); // "//" の先頭へ移動
            // Editor.Delete(2); // 2文字削除 <-- この行を変更
            Editor.Delete(); // 1文字目 (/) を削除
            Editor.Delete(); // 2文字目 (/) を削除
        }
        
        Editor.Down(0); // 次の行へ
    }
}