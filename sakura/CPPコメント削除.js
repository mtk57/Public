// C++のコメントを削除するsakuraエディタマクロ

// 全テキストを取得
SelectAll(0);
var text = GetSelectedString(0);
GoFileTop();

// コメントを削除する関数
function removeComments(str) {
    var result = '';
    var i = 0;
    var len = str.length;
    var inString = false;
    var stringChar = '';
    
    while (i < len) {
        var ch = str.charAt(i);
        var nextCh = (i + 1 < len) ? str.charAt(i + 1) : '';
        
        // 文字列リテラル内の処理
        if (inString) {
            result += ch;
            // エスケープ文字の処理
            if (ch === '\\' && i + 1 < len) {
                result += nextCh;
                i += 2;
                continue;
            }
            // 文字列の終了
            if (ch === stringChar) {
                inString = false;
            }
            i++;
            continue;
        }
        
        // 文字列リテラルの開始
        if (ch === '"' || ch === "'") {
            inString = true;
            stringChar = ch;
            result += ch;
            i++;
            continue;
        }
        
        // 単行コメント //
        if (ch === '/' && nextCh === '/') {
            // 行末まで読み飛ばす
            i += 2;
            while (i < len && str.charAt(i) !== '\n' && str.charAt(i) !== '\r') {
                i++;
            }
            // 改行は保持
            if (i < len && (str.charAt(i) === '\n' || str.charAt(i) === '\r')) {
                result += str.charAt(i);
                // \r\nの場合は両方処理
                if (i + 1 < len && str.charAt(i) === '\r' && str.charAt(i + 1) === '\n') {
                    i++;
                    result += str.charAt(i);
                }
                i++;
            }
            continue;
        }
        
        // 複数行コメント /* */
        if (ch === '/' && nextCh === '*') {
            // */が見つかるまで読み飛ばす
            i += 2;
            while (i < len - 1) {
                if (str.charAt(i) === '*' && str.charAt(i + 1) === '/') {
                    i += 2;
                    break;
                }
                i++;
            }
            continue;
        }
        
        // 通常の文字
        result += ch;
        i++;
    }
    
    return result;
}

// コメントを削除
var cleanedText = removeComments(text);

// テキストを置換
SelectAll(0);
InsText(cleanedText);
GoFileTop();

// 完了メッセージ
MessageBox("C++のコメントを削除しました。");