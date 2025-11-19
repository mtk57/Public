// [マクロ] 選択行の "//@ " 以降を削除（修正版）
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
    
    // 最初の "//" の位置を探す
    var nFirstCommentPos = sLine.indexOf("//");
    
    if (nFirstCommentPos >= 0) {
        // 最初の "//" が "//@ " かどうかをチェック
        if (sLine.substr(nFirstCommentPos, 4) == "//@ ") {
            // "//@ " より前の部分を保存
            var sBeforeComment = sLine.substring(0, nFirstCommentPos);
            
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
}

/* 以下はテストデータと期待値

// テストデータ - マクロ実行前のサンプルコード

  HogeFunc(); //@ hogehoge
  var x = 10; //@ これは変数
  FugaFunc(); //@ 関数呼び出し
  var y = 20;
console.log(x); //@ ログ出力
    var z = 30; //@ インデントあり
//@ 行頭コメント
  // 通常のコメント
  BarFunc(); // 通常のコメント（//@ ではない）
  BazFunc();//@ スペースなしコメント
  QuxFunc();  //@ 前に複数スペース
  MultiFunc(); //@ コメント1 //@ コメント2



// テストデータ - マクロ実行後の期待結果

  HogeFunc();
  var x = 10;
  FugaFunc();
  var y = 20;
console.log(x);
    var z = 30;

  // 通常のコメント
  BarFunc(); // 通常のコメント（//@ ではない）
  BazFunc();//@ スペースなしコメント
  QuxFunc();
  MultiFunc();

*/


