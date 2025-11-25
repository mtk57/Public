// コメント位置を揃える ver1.0.0
(function(){
    var TAB_WIDTH = 4;
    var sel = Editor.GetSelectedString(0);
    if (!sel || sel.length === 0) {
        Editor.MessageBox("範囲を選択してください。", 0);
        return;
    }
    
    // 改行コードを取得
    var lineCodes = ['\r\n', '\r', '\n'];
    var codeIndex = Editor.GetLineCode();
    var lc = (typeof codeIndex === 'number' && codeIndex >= 0 && codeIndex <= 2) ? lineCodes[codeIndex] : '\n';
    
    var lines = sel.split(/\r\n|\r|\n/);
    
    // 視覚幅を計算する関数
    function visualWidth(s) {
        var w = 0;
        for (var i = 0; i < s.length; i++) {
            var ch = s.charAt(i);
            var code = s.charCodeAt(i);
            if (ch === '\t') {
                w += TAB_WIDTH;
            } else if (code <= 0x007f) {
                w += 1;
            } else if (code >= 0xFF61 && code <= 0xFF9F) {
                w += 1;
            } else {
                w += 2;
            }
        }
        return w;
    }
    
    // 指定回数文字を繰り返す
    function repeatChar(count, ch) {
        if (count <= 0) return '';
        return Array(count + 1).join(ch);
    }
    
    // 各行を解析: "//"の位置を見つける
    var lineData = [];
    var maxCommentPos = 0;
    
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var commentIndex = line.indexOf('//');
        
        if (commentIndex === -1) {
            // コメントがない行はそのまま
            lineData.push({
                original: line,
                hasComment: false
            });
        } else {
            // "//"より前の部分のタブを半角スペース4つに置換
            var beforeComment = line.substring(0, commentIndex);
            beforeComment = beforeComment.replace(/\t/g, repeatChar(TAB_WIDTH, ' '));
            
            var comment = line.substring(commentIndex);
            var width = visualWidth(beforeComment);
            
            lineData.push({
                original: line,
                hasComment: true,
                beforeComment: beforeComment,
                comment: comment,
                width: width
            });
            
            if (width > maxCommentPos) {
                maxCommentPos = width;
            }
        }
    }
    
    // 各行を再構築
    var outLines = [];
    for (var j = 0; j < lineData.length; j++) {
        var data = lineData[j];
        
        if (!data.hasComment) {
            outLines.push(data.original);
        } else {
            var padding = maxCommentPos - data.width;
            var newLine = data.beforeComment + repeatChar(padding, ' ') + data.comment;
            outLines.push(newLine);
        }
    }
    
    var outText = outLines.join(lc);
    Editor.InsText(outText);
})();