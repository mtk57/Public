// sakuraエディタ用 タグジャンプマクロ
// SimpleMethodCallListCreator.exe を呼び出してタグジャンプを実行する

// TODO: 利用環境に合わせて実行ファイルパスを設定してください。
var SIMPLE_METHOD_CALL_LIST_CREATOR_EXE = "C:\\\\Path\\\\To\\\\SimpleMethodCallListCreator.exe";

var TAG_PREFIX = "//@ ";

function main() {
    if (!Editor) {
        alert("Editor オブジェクトが利用できません。");
        return;
    }

    if (!SIMPLE_METHOD_CALL_LIST_CREATOR_EXE || SIMPLE_METHOD_CALL_LIST_CREATOR_EXE.length === 0) {
        Editor.Inform("SimpleMethodCallListCreator.exe のパスが設定されていません。");
        return;
    }

    var lineNo = Editor.GetCaretLine();
    if (lineNo < 0) {
        Editor.Inform("カーソル位置を取得できませんでした。");
        return;
    }

    var lineText = Editor.GetLineStr(lineNo);
    if (!lineText) {
        Editor.Inform("現在行にテキストがありません。");
        return;
    }

    var prefixIndex = lineText.indexOf(TAG_PREFIX);
    if (prefixIndex < 0) {
        Editor.Inform("行内にタグジャンプ情報が見つかりません。");
        return;
    }

    var tagContent = lineText.substring(prefixIndex + TAG_PREFIX.length).trim();
    if (!tagContent) {
        Editor.Inform("タグジャンプ情報が空です。");
        return;
    }

    var separatorIndex = tagContent.indexOf("\t");
    if (separatorIndex < 0) {
        Editor.Inform("タグジャンプ情報の形式が不正です。");
        return;
    }

    var filePath = tagContent.substring(0, separatorIndex).trim();
    var methodSignature = tagContent.substring(separatorIndex + 1).trim();

    if (!filePath) {
        Editor.Inform("タグジャンプ情報からファイルパスを取得できませんでした。");
        return;
    }

    if (!methodSignature) {
        Editor.Inform("タグジャンプ情報からメソッドシグネチャを取得できませんでした。");
        return;
    }

    launchTagJump(filePath, methodSignature);
}

function launchTagJump(filePath, methodSignature) {
    try {
        var shell = new ActiveXObject("WScript.Shell");
        var command = buildCommandLine(
            SIMPLE_METHOD_CALL_LIST_CREATOR_EXE,
            filePath,
            methodSignature
        );
        shell.Run(command, 0, false);
    } catch (ex) {
        Editor.Inform("SimpleMethodCallListCreator.exe の起動に失敗しました。\n" + ex.message);
    }
}

function buildCommandLine(exePath, filePath, methodSignature) {
    return [
        quoteForCmd(exePath),
        quoteForCmd(filePath),
        quoteForCmd(methodSignature)
    ].join(" ");
}

function quoteForCmd(value) {
    if (value === null || value === undefined) {
        return "\"\"";
    }

    var escaped = String(value).replace(/"/g, "\"\"");
    return "\"" + escaped + "\"";
}

main();
