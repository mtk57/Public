## 出力言語
このリポジトリに関するすべての説明・要約は日本語で記述してください。

## SimpleOCRforBase32 仕様  Version 1.1.0
===============

## 検討履歴メモ (Codex)
- OCR は Base32 (32桁) + ハイフン + チェックサム(2桁) の行を対象。誤認識が約 12 行あり、精度向上策を検討した。
- 現行ロジックは OCR 結果からチャンク32桁とチェックサム2桁を取り出し、自前の `ComputeChecksum` で一致判定し重み付けに利用しているが、チェックサム自体の誤認識がある場合は過信しない方針が必要。
- 改善案: 画像前処理の閾値・リサイズ・モルフォジー調整、Tesseract の PSM/ホワイトリスト/辞書設定の見直し、N-gram 等での事後補正強化、誤行ログ分析、フォント特化の補助 OCR 等。
- UI 改善案: TextBox から `DataGridView` へ変更し、各行のチャンク/チェックサム/一致フラグを表示。OCR 後に不一致行のみユーザーがチェックサムやチャンクを編集→再調整を実行する操作フロー。
- チェックサムを人間が確認・入力するワークフローも検討。特に「一致しない行だけ手入力→再調整」方式が実用的。
- Base32 に存在しない文字（例: `1`, `8`）が最終結果に残らないよう、投票段階や出力段階で `AllowedCharacterSet` による強制フィルタと UI の入力制限を追加する案。
- `PreprocessImage` 内で OpenCV の deskew を実行しているが、角度推定の安定化（二値化後に `FindNonZero` する等）と回転による欠け防止（ボーダー処理）が改善余地として挙がった。

## 実装概要
- WinForms (`Form1`) で構成。画像パス指定（テキストボックス／ドラッグ＆ドロップ／ファイルダイアログ）と結果表示テキストボックス、並びに前処理調整ボタンを備える。
- `ImageEditForm` はコントラスト・明るさ・二値閾値のスライダーでプレビューし、保存すると別ファイル（`_edit###`）を生成してメイン画面に戻す。
- OCR 実行前に `tessdata` ディレクトリ存在を検証し、OpenCV での前処理結果（deskew→グレースケール→ガウシアンぼかし→適応的二値化→クロージング→1.6倍リサイズ）と、行ごとの推定矩形を一時PNGとして用意する。
- `PreprocessImage` では行スキャンし、非ゼロ画素の連続領域から `OpenCvSharp.Rect` を生成し、後の行単位OCRに渡す。中間画像（deskew/binary/denoise）はデバッグ用にアプリ基準ディレクトリへ保存。

## OCR 処理フロー
- `RunOcr` は前処理済み画像と元画像の両方を `ProcessingSource` として扱い、各ソースに対して Tesseract の `classify_bln_numeric_mode` を true/false の2パターンで走らせる。ページ分割モードは列単位 (`SingleColumn`) と行矩形 (`SingleLine`) を使い分ける。
- エンジンは `tessedit_char_whitelist` に Base32 文字＋ハイフンのみを許可し、辞書系 (`load_system_dawg` 等) を無効化している。`AllowedCharacterSet` と `SoftEquivalents(0→O,1→I,8→B)` により無効文字を排除し、32桁未満の行は候補から除外。
- OCR結果は `CandidateResult` として正規化文字列、信頼度、モード種別を保持し、各行を32桁チャンク＋2桁チェックサムとして `CandidateLineInfo` に分解。重みは `confidence + validLines*0.001 + checksumMatches*0.01` で計算する。
- `MergeCandidates` は行番号ごとに候補の `CharVote` を集計し、チェックサム一致行を1.25倍で加点して多数決。`AmbiguityMap` と `ApplyBigramAdjustments` で類似文字の出力を安定化し、足りない位置は重み順にフォールバック。
- `ComputeChecksum` は 33x XOR ハッシュから 10bit を取り出し Base32 アルファベットに写像して 2文字生成。投票結果と既存候補のハミング距離が4以下なら候補行のチャンクを採用し、最終行は `xxxx xxxx-YY` 形式で整形して連結。
- 目標チェックサムは候補の重み投票で決定し、`TryAdjustChunkWithChecksum` が曖昧位置（最大5箇所）を探索して合わせ込みを試行。`ApplyChecksumPreference` や `DetermineTargetChecksum` によりチェックサム一致情報を重視した再調整を行う。

## 補助機能・留意点
- `BtnImageAdjust` から `ImageEditForm` を開き、調整後画像が OK で閉じられた場合は新しいパスをメインフォームへ戻す。
- `SaveDeskewedImage` などの補助出力は失敗しても例外を握りつぶし、メインフローを阻害しない設計。
- `btnStart` クリック時は非同期で OCR を実行し、完了までボタンを無効化。例外メッセージは MessageBox でユーザーに提示。
