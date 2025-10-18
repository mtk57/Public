// 改行コードを揃える  ver1.0.0
(function(){
    var TAB_WIDTH = 4;

    var sel = Editor.GetSelectedString(0);
    if (!sel || sel.length === 0) {
        Editor.MessageBox("範囲を選択してください。", 0);
        return;
    }

    // 安全に改行コードを決める
    var lineCodes = ['\r\n', '\r', '\n'];
    var codeIndex = Editor.GetLineCode();
    var lc = (typeof codeIndex === 'number' && codeIndex >= 0 && codeIndex <= 2) ? lineCodes[codeIndex] : '\n';

    var lines = sel.split(/\r\n|\r|\n/);

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

    // repeat の代わりに Array.join を使う helper
    function repeatChar(count, ch) {
        if (count <= 0) return '';
        // Array(count+1).join(ch) は count 個の ch を作る（空要素間に ch を挿むため）
        return Array(count + 1).join(ch);
    }

    var maxW = 0;
    for (var i = 0; i < lines.length; i++) {
        var w = visualWidth(lines[i]);
        if (w > maxW) maxW = w;
    }

    var outLines = [];
    for (var j = 0; j < lines.length; j++) {
        var cur = lines[j];
        var curW = visualWidth(cur);
        var pad = maxW - curW;
        if (pad > 0) {
            cur = cur + repeatChar(pad, ' ');
        }
        outLines.push(cur);
    }

    var outText = outLines.join(lc);
    Editor.InsText(outText);
})();
