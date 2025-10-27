## 出力言語
このリポジトリに関するすべての説明・要約は日本語で記述してください。

## SimpleSqlAdjuster 仕様  Version 1.0.0

## アプリ概要
- Windows Forms (.NET Framework 4.8) 製のデスクトップアプリ。エントリポイントは `Program.cs` の `MainForm` 起動。
- 1行につき1 SQL を前提に、テキストボックスに貼り付けた SQL を整形し、変数名とフォーマット済み SQL を出力する。
- 設定 (`SimpleSqlAdjuster.settings.json`) とログ (`SimpleSqlAdjuster.log`) は実行ファイルと同じディレクトリに保存する。

## UI 構成（`MainForm`）
- `txtBeforeSQL`（入力）と `txtAfterSQL`（出力）の 2 つのマルチラインテキストボックス、`btnRun` ボタンのみで構成。
- フォームロード時に `UserSettingsService` を通じてウィンドウサイズ・位置・前回入力を復元。閉じる直前に同情報を保存する。
- 実行ボタンクリックで `SqlAdjuster.Process` を呼び出し、成功時は整形結果を表示、エラー時はメッセージボックスと `txtAfterSQL` にエラーメッセージを表示し、ログへ出力する。
- 出力テキストボックスは読み取り専用で、自動スクロールにより最新行へフォーカスさせる。

## SQL 整形パイプライン
1. `SqlAdjuster.Process` が入力文字列を行単位で読み込み、空行はスキップ。結果は SQL ごとに空行で区切って連結。
2. 行の先頭に変数宣言（`変数名=SQL`）がある場合は左辺を変数名として切り出し、右辺の SQL を処理。コメント行 (`--`, `/*`) はそのまま返す。
3. SQL の開始キーワードを `SELECT`, `WITH`, `UPDATE` など既知の語から判定し、見つからない場合は `SqlProcessingException` を送出。
4. `MacroExpander` が `:_WHERE_AND(...)` / `:_WHERE_OR(...)` マクロを解析し、ドット区切りの条件を `WHERE` 句に展開。入れ子の括弧や空引数も検証してエラー報告する。
5. `SqlFormatter.Format` が SQL をトークン化し、句単位で改行・インデント（スペース 2 文字）を整形。コメント・リテラル・パラメータ・Oracle の旧構文記号 (`(+)` 等) に対応。
6. 変数名が存在する場合は 1 行目に変数名、2 行目以降に整形済み SQL を配置する。

## `SqlFormatter` の構成
- `SqlTokenizer` がキーワード・識別子・記号・コメント・文字列・数値・パラメータ（`:` / `@` 始まり）を切り出し、Oracle 由来のキーワードもカバーする。
- `StatementFormatter` がトークン列を解析。`SELECT`/`FROM`/`WHERE`/`GROUP BY`/`ORDER BY`/`HAVING`/`UNION` など主要句を認識し、句ごとに `SqlWriter` へ出力。
- `SqlWriter` はインデントレベルをスペース 2 文字単位で反映し、行末の余分な空白を除去した状態で文字列を構築する。
- `WHERE` 句は AND/OR ごとに行を分割し、`FROM` 句は JOIN 種別やカンマ区切りで改行を入れるなど、読みやすさ重視のルールを実装。

## エラー処理・ロギング
- 整形中に発生した問題は `SqlProcessingException` として捕捉し、行・列番号を保持。`ToDisplayMessage()` で「行 X, 列 Y: メッセージ」の形式に整形する。
- 例外やアプリ内メッセージは `LogService` により `SimpleSqlAdjuster.log` に追記。ログ書き込み失敗時はアプリ動作を妨げない。

## 設定ファイル
- `UserSettingsService` が `JavaScriptSerializer` を用いて JSON をシリアライズ／デシリアライズ。
- 保存項目はウィンドウ幅・高さ・位置 (`WindowWidth`, `WindowHeight`, `WindowX`, `WindowY`) と直前に入力した SQL (`LastBeforeSql`)。
- 設定ファイルが存在しない、または読み書き例外が発生した場合はログへ出力し、アプリはデフォルト値で継続する。

## ビルド／実行補足
- ソリューションファイル `SimpleSqlAdjuster.sln` とプロジェクトファイル `SimpleSqlAdjuster.csproj` を Visual Studio などで開けばビルド可能。
- 実行バイナリと同じフォルダに設定・ログが生成されるため、配布時は書き込み可能な場所に配置する。
