// [マクロ] 選択行の "//@ " 以降を削除
var nLineFrom = Editor.GetSelectLineFrom();
var nLineTo = Editor.GetSelectLineTo();

// 選択範囲がない場合はカレント行のみ処理
if (nLineFrom == 0 || nLineTo == 0) {
    var nCurrentLine = Editor.ExpandParameter("$y");
    nLineFrom = parseInt(nCurrentLine);
    nLineTo = parseInt(nCurrentLine);
}

for (var i = nLineFrom; i <= nLineTo; i++) {
    Editor.MoveCursor(i, 1, 0); // 行頭に移動
    var sLine = Editor.GetLineStr(0); // 現在行の文字列取得（改行コード含む）
    
    // "//@ " を探索
    var nCommentPos = sLine.indexOf("//@ ");
    
    if (nCommentPos >= 0) {
        // "//@ " より前の部分を保存
        var sBeforeComment = sLine.substring(0, nCommentPos);
        
        // 末尾の空白を削除
        sBeforeComment = sBeforeComment.replace(/\s+$/, '');
        
        // 改行コードを抽出（元の行から）
        var sLineBreak = "";
        if (sLine.match(/\r\n$/)) {
            sLineBreak = "\r\n";
        } else if (sLine.match(/\n$/)) {
            sLineBreak = "\n";
        } else if (sLine.match(/\r$/)) {
            sLineBreak = "\r";
        }
        
        // 行全体を選択
        Editor.MoveCursor(i, 1, 0);
        Editor.SelectLine(0);
        
        // 新しい内容に置換（改行コード付き）
        Editor.InsText(sBeforeComment + sLineBreak);
    }
}