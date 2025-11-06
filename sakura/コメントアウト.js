// [マクロ] 選択行を一括でコメントアウト (//) - 修正版 v3

// 選択行の開始行と終了行を取得
var nLineFrom = Editor.GetSelectLineFrom();
var nLineTo = Editor.GetSelectLineTo();

// 選択範囲がない場合は何もしない
if (nLineFrom == 0 || nLineTo == 0) {
    // 処理終了
} else {
    // カーソルを選択開始行の先頭に移動 (これで選択も解除されます)
    Editor.MoveCursor(nLineFrom, 1, 0); // 1は行頭(桁)、0は非選択

    // ループ処理：開始行から終了行まで
    for (var i = nLineFrom; i <= nLineTo; i++) {
        Editor.GoLineTop(0); // 論理行頭へ移動 (インデントを無視して行の先頭へ)
        // Editor.InsertString("//"); // "//"を挿入  <-- この行を変更
        Editor.InsText("//"); // "//"を挿入
        Editor.Down(0); // 次の行へ移動
    }
}