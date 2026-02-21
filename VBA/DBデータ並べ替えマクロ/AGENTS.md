# AsIs2ToBe DBデータ並べ替えマクロ

■背景（実現したいこと）
AsIsのDBデータがExcelファイルで提供されている。
このDBデータをToBeのDBにCSVインポートしたいが、AsIsとToBeではテーブル名、カラム名などが
一部不一致のものがあるため、そのままで使用できない。
AsIsのDBデータをToBeのスキーマ定義に合わせて並び替える作業を自動化したい。

<DBデータのイメージ>
AsIs.xlsx                                 ToBe.xlsx
[HOGE]シート                           →   [HOGEHOGE]シート
C1 C2 C3 C4 C5                          C1 C5 C6 C3 C4 C7
A  B  C  D  E                           A  E     C  D  E
V  W  X  Y  Z                           V  Z     X  Y  1

<ToBe/AsIsマッピング定義のイメージ>（ToBe基準）
ToBeAsIsMapping.xlsx
[mapping]シート
ToBeテーブル名  ToBeカラム名   AsIsテーブル名  AsIsカラム名
HOGEHOGE      C1            HOGE          C1
HOGEHOGE      C5            HOGE          C5
HOGEHOGE      C6
HOGEHOGE      C3            HOGE          C3
HOGEHOGE      C4            HOGE          C4
HOGEHOGE      C7


■ツール仕様（現時点実装）

① ツールはExcel VBAマクロで作成する。（便宜上、ツール.xlsmと呼ぶ）

② ツール.xlsmには以下のシートがある。
- [mapping]
- [target]
- [macro]
- [log]（ログ出力用）

③ [mapping]シート入力項目（開始行は5行目）
- B5〜Bxx: ToBeのテーブル名（物理名） 必須
- C5〜Cxx: ToBeのテーブル名（論理名） 任意
- D5〜Dxx: ToBeのカラム名（物理名） 必須
- E5〜Exx: ToBeのカラム名（論理名） 任意
- F5〜Fxx: ToBeのキー（主キー:X、ユニークキー:U） 任意
- G5〜Gxx: 空
- H5〜Hxx: ToBeのデータ型 任意
- I5〜Ixx: ToBeの桁数 任意
- J5〜Jxx: ToBeの桁数(小数部) 任意
- K5〜Kxx: 空
- L5〜Lxx: ToBeの説明 任意
- M5〜Mxx: ToBeの備考 任意
- N5〜Nxx: 空
- O5〜Oxx: AsIsのテーブル名（物理名） 必須（AsIs側が存在する列の場合）
- P5〜Pxx: AsIsのテーブル名（論理名） 任意
- Q5〜Qxx: AsIsのカラム名（物理名） 必須（AsIs側が存在する列の場合）
- R5〜Rxx: AsIsのカラム名（論理名） 任意

④ [target]シート入力項目（開始行は5行目）
- C5〜Cxx: AsIsデータのExcelファイルパス（絶対パス） 必須
- D5〜Dxx: AsIsデータのExcelファイルのシート名 必須
- E5〜Exx: AsIsのテーブル名（物理名） 必須
- F5〜Fxx: AsIsのカラム名（物理名）の開始セル（A1形式） 必須
- G5〜Gxx: AsIsのデータの開始セル（A1形式） 必須

⑤ [macro]シートのボタン
- 並び替え開始
- ToBeマッピング元ネタ作成支援

⑥ 並び替え開始ボタン押下時の処理概要
a) 実行確認メッセージ（Yes/No）を表示する。
- 「並び替えを実行しますか?」

b) 入力チェックを行う（mapping/target）。
- エラー時は [log] シートに出力し、メッセージボックスを表示して全体中断。

c) 実処理前に target の全行を事前チェックする。
- ファイル存在チェック
- シート存在チェック
- mappingで必要なAsIsカラム存在チェック（ヘッダ比較は大文字小文字を区別しない）
- 出力先ファイル（末尾 `_TOBE`）の既存チェック

d) targetの元ファイルをコピーし、コピー先を処理対象とする。
- コピー先ファイル名は「拡張子の前に `_TOBE` を付与」
- 同一元ファイルがtargetに複数行ある場合、コピーは1回のみ

e) コピー先ファイル内の対象シートを mapping に従って再構成する。
- ToBe列順で並び替え
- ToBeにしかない列は空で追加
- mappingに必要なAsIsカラムが無ければエラー

f) 同一コピー先ファイルはまとめて処理し、最後に保存する。

⑦ ToBeマッピング元ネタ作成支援ボタン押下時の処理概要
[macro]シートの入力項目:
- B20: ToBeテーブル定義ファイルパス（絶対パス）
- B21: テーブル物理名セル位置（A1形式）
- B22: テーブル論理名セル位置（A1形式）
- B23: ToBeカラム名（物理名）セル位置（A1形式）
- B24: ToBeカラム名（論理名）セル位置（A1形式）
- B25: ToBeキーセル位置（A1形式）
- B26: ToBeデータ型セル位置（A1形式）
- B27: ToBe桁数セル位置（A1形式）
- B28: ToBe桁数(小数部)セル位置（A1形式）
- B29: ToBe説明セル位置（A1形式）
- B30: ToBe備考セル位置（A1形式）
- B31: ToBeカラム削除指定セル位置（A1形式）
- B32〜Bxx: ToBeテーブル定義ファイルのシート名

処理:
a) B20のファイルを開き、B32以降の指定シートを順に処理する。
b) 新規シート（`mapping_seed_yyyymmdd_hhnnss`）を作成し、mappingシートの1〜4行目をコピーする。
c) 各シートで B23（ToBeカラム物理名）を起点に下方向へ走査し、空セルになるまで転記する。
d) B31（削除指定）が `X` の行は転記対象外。
e) 転記先は mapping と同じ列位置（B,C,D,E,F,H,I,J,L,M）。
f) AsIs側列（O,Qなど）は空のまま。

g) エラー時は [log] に出力し中断、完了時は件数を [log] とメッセージで通知する。

⑧ 主な入力不正チェック（実装済み）
- mapping: 同一ToBeテーブル内のToBeカラム重複はエラー
- mapping: AsIsテーブルが複数ToBeテーブルへ紐づく定義はエラー
- target: 絶対パス/A1形式不正はエラー
- target: データ開始セル行はカラム名開始セル行より下であること
