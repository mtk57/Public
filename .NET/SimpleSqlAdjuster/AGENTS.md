## 出力言語
このリポジトリに関するすべての説明・要約は日本語で記述してください。

## SimpleSqlAdjuster 仕様  Version 1.0.0
===============

#概要
----
- Windows Forms (.NET Framework 4.8) ベースのデスクトップ アプリケーション。
- 画面のテキストボックスに張られたSQL文字列を見やすく整形する
- ユーザー設定は JSON 形式の設定ファイルに保存される。(EXEと同じフォルダに保存)

#画面
----
①変更前SQL：TextBox：整形前のSQLを入力する。（※1）
②変更後SQL：TextBox：整形後のSQLが出力される。（※2）
③実行：Button：クリックすることで①の内容を解析・整形し、②に出力する。

※1
入力されるSQLは以下のようになっている。
----
・1行につき1つのSQL

・SQLが始まる前に変数のような文字列がある。（いきなりSQLから始まっていても解析可能とする）
例1：
Hoge1=SELECT * FROM FUGA
→SQLの前にある変数のような文字列の例。必ず半角"="があり、左辺が変数、右辺はSQL本体となる。

・このような行を複数行入力しても解析可能とする。

・SQLはOracleのSQLの古い構文も解析可能とする。（joinとかに(+)とか使ったりする）

・where句には以下のようなパラメータが指定されることもある。
例2：
WHERE HOGE = :Hoge

・where句には以下のようなマクロが指定されることもある。
例3：
:_WHERE_AND(HOGE = :hoge. FUGA = :fuga)
→これは以下のWHERE句として解析する
WHERE HOGE = :hoge AND FUGA = :fuga

・where句には以下のようなマクロが指定されることもある。
例4：
:_WHERE_OR(HOGE = :hoge. FUGA = :fuga)
→これは以下のWHERE句として解析する
WHERE HOGE = :hoge OR FUGA = :fuga
----


※2
出力するSQLの形式は以下とする。
1行目：変数があった場合は変数名（"="は不要）
2行目以降：整形後のSQL

SQLの整形について
・以下のようにインデントは2として、タブ文字は使わない。

例A：
SELECT
  HOGE_CLM1,
  FUGA_CLM2
FROM
  MYTABLE
WHERE
  HOGE_CLM1 = :hoge
  AND
  FUGA_CLM2 = :fuga

例B：
SELECT
  HOGE_CLM1,
  FUGA_CLM2
FROM
  MYTABLE
:_WHERE_AND
(
  HOGE_CLM1 = :hoge,
  FUGA_CLM2 = :fuga
)


----

#エラー処理
----
- 解析時にエラーが発生した場合は行・列・原因などをログファイルと画面に出力する。
- ログファイルはEXEと同じフォルダに保存する。

