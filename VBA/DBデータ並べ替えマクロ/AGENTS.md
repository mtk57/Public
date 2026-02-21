# AsIs2ToBe DBデータ並べ替えマクロ

	■背景（実現したいこと）
		AsIsのDBデータがExcelファイルで提供されている。
		このDBデータをToBeのDBにCSVインポートしたいが、AsIsとToBeではテーブル名、カラム名などが
		一部不一致のものがあるため、そのままで使用できない。
		AsIsのDBデータをToBeのスキーマ定義に合わせて並び替える作業を自動化したい。

		<DBデータのイメージ>
			AsIs.xlsx								ToBe.xlsx
			[HOGE]シート						→		[HOGEHOGE]シート
			C1	C2	C3	C4	C5		→		C1	C5	C6	C3	C4	C7
			A	B	C	D	E		→		A	E		C	D	E
			V	W	X	Y	Z		→		V	Z		X	Y	1

		<ToBe/AsIsマッピング定義のイメージ>										※ToBe基準
			ToBeAsIsMapping.xlsx							ToBe		AsIs
			[mapping]シート																		説明
			テーブル名				カラム名					テーブル名				カラム名
			HOGEHOGE				C1					HOGE				C1					ToBeとAsIsはカラム名、位置が全て一致（並び替え不要）
			HOGEHOGE				C5					HOGE				C5					ToBeとAsIsでカラム名は同じだが位置が変わった（並び替えが必要）
			HOGEHOGE				C6														ToBeで新規追加のカラムなのでマッピング不要
			HOGEHOGE				C3					HOGE				C3					ToBeとAsIsでカラム名は同じだが位置が変わった（並び替えが必要）
			HOGEHOGE				C4					HOGE				C4					ToBeとAsIsでカラム名は同じだが位置が変わった（並び替えが必要）
			HOGEHOGE				C7														ToBeで新規追加のカラムなのでマッピング不要
																					AsIsのC2はToBeでは削除されたのでマッピング不要


	■ツール仕様
		①	ツールはExcel VBAマクロで作成する。（便宜上、ツール.xlsmと呼ぶ）
		②	ツール.xlsmには以下のシートがある。
			[mapping]
			[target]
			[macro]
		③	[mapping]シートに以下の入力項目がある。
				B2〜Bxx			ToBeのテーブル名（物理名）												必須
				C2〜Cxx			ToBeのテーブル名（論理名）												任意
				D2〜Dxx			ToBeのカラム名（物理名）												必須		※ToBeのテーブル名（物理名）が空文字以外の場合のみ必須。以降の必須項目も同じ
				E2〜Exx			ToBeのカラム名（論理名）												任意
				F2〜Fxx			ToBeのキー（主キー：X、ユニークキー：U）												任意
				G2〜Gxx			空												-
				H2〜Hxx			ToBeのデータ型												任意
				I2〜Ixx			ToBeの桁数												任意
				J2〜Jxx			ToBeの桁数(小数部)												任意
				K2〜Kxx			空												-
				L2〜Lxx			ToBeの説明												任意
				M2〜Mxx			ToBeの備考												任意
				N2〜Nxx			空												-
				O2〜Oxx			AsIsのテーブル名（物理名）												必須
				P2〜Pxx			AsIsのテーブル名（論理名）												任意
				Q2〜Qxx			AsIsのカラム名（物理名）												必須
				R2〜Rxx			AsIsのカラム名（論理名）												任意
		④	[target]シートに以下の入力項目がある。
				C2〜Cxx			AsIsデータのExcelファイルパス（絶対パス）												必須
				D2〜Dxx			AsIsデータのExcelファイルのシート名												必須		※AsIsデータのExcelファイルパス（絶対パス）が空文字以外の場合のみ必須。以降の必須項目も同じ
				E2〜Exx			AsIsのテーブル名（物理名）												必須
				F2〜Fxx			AsIsのカラム名（物理名）の開始セル（A1形式）												必須
				G2〜Gxx			AsIsのデータの開始セル（A1形式）												必須
		⑤	[macro]シートに以下のボタンがある。
				並び替え開始
		⑥	並び替え開始ボタン押下時の処理概要
				a)	必須入力項目チェック
					→エラー時はどこが間違っているかが分るメッセージを[macro]シートに出力して処理を中止する。（メッセージボックスも出す）
				b)	ファイル存在チェック
				c)	シート存在チェック
				d)	[target]シートのExcelファイルを開く。
				e)	[target]シートのシートをリネームコピーする。
					→同名シートがすでにある場合はエラー。リネームはシート名の末尾に"_ToBe"を付けるだけ。
				f)	[mapping]シートに従って列を並び替える。ToBeにしか無い列がある場合を追加する。（データは空）
					→[mapping]シートにあるべき列が無い場合はエラー。
				g)	d)～f)を繰り返す。
