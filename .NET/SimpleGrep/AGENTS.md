## 出力言語
このリポジトリに関するすべての説明・要約は日本語で記述してください。

## SimpleGrep 仕様  Version 1.4.6
===============

概要
----
- Windows Forms (.NET Framework 4.8) ベースのデスクトップ アプリケーションで、複数ファイルに対する grep 風検索を提供します。
- エントリポイントの `Program.cs` は `MainForm` を起動し、すべてのユーザー インターフェイスとロジックを内包します。
- ユーザー設定は JSON 形式の設定ファイルに保存され、必要に応じて桜エディタとの連携（ナビゲーション／エクスポート）を行います。

ユーザー インターフェイス
----------------
- 検索対象フォルダー (`cmbFolderPath`)、ファイル パターン (`comboBox1`、例: `*.cs`)、キーワード／正規表現 (`cmbKeyword`) の各コンボボックス。`cmbKeyword` は Enter キー押下で検索を実行する。
- チェックボックス:
  - `chkSearchSubDir`: サブディレクトリを検索対象に含める。
  - `chkUseRegex`: 正規表現検索を有効化する。
  - `chkCase`: 大文字小文字を区別する。
  - `chkTagJump`: ダブルクリック時に桜エディタのタグジャンプを使用する。
  - `chkMethod`: メソッド導出の有無。DataGridViewの「メソッド」列にメソッドシグネチャ(戻り値 メソッド名(引数))を出力するか否かを指定する。現在はJava(.java)とC#(.cs)をサポートする。
- ボタン:
  - `btnBrowse`: `FolderBrowserDialog` を開き、検索ルートを選択する。
  - `button1`: 検索を実行する（`btnGrep_Click` にバインド）。
  - `btnCancel`: 実行中の検索をキャンセルする（未実行時は無効化）。
  - `btnExportSakura`: 結果を `.grep` ファイルにエクスポートし、桜エディタが利用可能な場合は開く。
- `dataGridViewResults`: ファイル パス、行番号、該当行テキスト、および実行時に追加される非表示のエンコーディング列を表示。「メソッド」列にはメソッドシグネチャを表示(メソッド導出チェックボックスがONの場合)
- 進捗表示: `progressBar`、`lblPer`（割合）、`labelTime`（経過時間）。

検索ワークフロー
----------------
- `btnGrep_Click` により検索を開始。
- 検索開始時にキャンセルトークンを生成し、進行中は `btnCancel` で中止要求を送出できる。
- フォルダー、ファイル パターン、検索パターンを検証し、無効な入力があればユーザーに通知。
- コンボボックスの履歴（各フィールド最大 10 件）を永続化。
- `chkSearchSubDir` がオンの場合は `Directory.GetFiles` と `SearchOption.AllDirectories` で対象ファイルを列挙。
- ファイル走査はバックグラウンド スレッド (`Task.Run`) 上で `Parallel.ForEach` を用いて実行:
  - BOM、UTF-8 判定ヒューリスティック、Shift_JIS を順に用いてエンコーディングを推定。
  - 推定したエンコーディングでファイルを行単位に読み込む。
  - `chkUseRegex` と `chkCase` の状態に応じて `Regex.IsMatch` または `string.IndexOf` を使用。
  - ファイル パス、行番号、行テキスト、エンコーディング名を格納した `ConcurrentBag<object[]>` に一致結果を蓄積。`chkMethod` がオンの場合、Java および C# ファイルについてメソッドシグネチャを解析し、「メソッド」列に表示する。
- 進捗報告は `IProgress<int>` を用い、100 ファイルごと、または完了時に通知。
- 完了後:
  - 収集した行を一括でグリッドにバインド。
  - UI（カーソルやボタン）を元に戻し、進捗バーを 100% にして経過時間（mm:ss）を更新。

結果の操作
----------
- ダブルクリック ハンドラー (`dataGridViewResults_CellDoubleClick`):
  - `chkTagJump` がオンの場合、`-Y={line}` 引数を使用して該当行で桜エディタを起動。
  - オフの場合はエクスプローラーで該当ファイルのフォルダーを開く。
  - ファイル欠如や起動失敗時にはメッセージ ボックスで通知。
- エクスポート ボタン (`btnExportSakura_Click`):
  - すべての結果を UTF-8 で `AppContext.BaseDirectory\yyyyMMdd_HHmmssfff.grep` に書き出す。
  - 各行のフォーマットは `{filePath}({line},1)  [{encoding}]: {content}`。
  - 桜エディタが起動できれば開き、失敗した場合は保存先を案内。

設定の永続化
------------
- 設定は実行ファイルと同じ場所の `SimpleGrep.settings.json` に保存。
- ロード時 (`MainForm_Load`):
  - 結果グリッドに非表示の「Encoding」列を追加。
  - `DataContractJsonSerializer` で履歴やチェックボックス状態を読み込み。
  - ウィンドウ タイトルにアセンブリ バージョン `ver {Major}.{Minor}.{Build}` を付与。
- 終了時 (`MainForm_FormClosing`): 現在の履歴とオプション状態をシリアル化して保存。

サクラエディタとの連携
-------------------
- `FindSakuraPath` の探索順序:
  1. `App.config` の `SakuraEditorPath` 値。
  2. `Program Files` / `Program Files (x86)` の既定インストール先。
  3. `PATH` に登録されたディレクトリ。
- サクラエディタが見つからない場合でも、エクスポート結果を保持したままユーザーに通知。



追加仕様  ※Version 1.4.2

- DataGridViewで選択した行（複数可）のファイルを「ファイルをコピー」ボタンで `Clipboard.SetFileDropList` に登録し、エクスプローラーのコピーと同等の状態にする。
  選択が無い、または実在しないパスのみの場合は情報メッセージを表示し、クリップボード操作に失敗した場合はエラーダイアログで通知する。

追加仕様  ※Version 1.4.4

- コメント無視チェックボックスがONの場合、Javaのコメント（//, /* 〜 */）にある検索キーワードは無視する。
- 将来的にJava以外の言語もサポートしやすいように設計すること。


