#Report Generator

##概要

Report Generator はエクセル形式の帳票出力プログラムを自動生成するためツールです。このツール及び、自動生成されるプログラムは VBScript ベースのため Windows 環境でのみ使用が可能です。IE 限定になりますが ActiveX の使用が有効な環境であれば Web アプリケーションに組み込むこともできます。

以下のような手順で作成します。

1. テンプレートの帳票エクセルファイルの作成
2. テンプレートファイルへのデータ識別子の記述
3. Report Generator による帳票出力プログラムの生成
4. アプリケーションへの組み込み

##チュートリアル

以下より Report Generator をダウンロードし解凍します。

[cyokodog / report-generator | GitHub](https://github.com/cyokodog/report-generator)

解凍すると下記構成のフォルダ、ファイルができあがります。([ ]はフォルダ)

	-[ReportGenerator]
		-ReportGenerator-*.*.*.wsf
		-[lib]
			-generateReport.vbs
			-init.vbs
			-jquery-1.4.2.min
		-[sample_app]
			-[ex01]
			-[ex02]
			・・・
		-[sample_template]
			-mitsumori.xls
			-mitsumori_map.xls
			-mitsumori_sample.xls

###テンプレートの帳票エクセルファイルの作成

まず、テンプレートとなる帳票エクセルファイルを作成します。sample_template フォルダの mitsumori.xls のように作ります。

**mitsumori.xls**

![mitsumori.xls](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921023813.png)

必要に応じ各セルに対し数式や書式等を設定します。

###テンプレートファイルへのデータ識別子の記述

可変データを埋め込む箇所に識別子を記述し、ファイル名の末尾に "_map" と付けてファイルを保存します。sample_template フォルダの mitsumori_map.xls のように記述します。

**mitsumori_map.xls**

![mitsumori_map.xls](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921023840.png)

識別子は {データ名} という形式で記述します。例えば納品期日なら {delivery_date} のようにして記述します。

###Report Generator による帳票出力プログラムの生成

mitsumori_map.xls を ReportGenerator フォルダの ReportGenerator-*.*.*.wsf に対しドロップします。

![drop to ...](http://cdn-ak.f.st-hatena.com/images/fotolife/c/cyokodog/20100921/20100921024637.png)

プログラム生成処理が始まるので、生成終了のメッセージが表示されるまで少し待ちます。

処理が完了するすると下記構成のフォルダ、ファイルができあがります。([ ]はフォルダ)

	-[result_mitsumori]
		-mitsumori.vbs
		-mitsumori.xls
		-mitsumori.xml
		-[bat_sample]
			-mitsumori.bat
			-mitsumori.wsf
			-mitsumori_dat.xml
		-[html_sample]
			-jquery-1.4.2.min.js
			-mitsumori.html
			-mitsumori_dat.xml

・  
・  
・  
・  
・  
・  
・  
疲れた・・・

続きは[旧バージョンのドキュメント](http://d.hatena.ne.jp/cyokodog/20100927/reportgenerator01)をご参照ください。基本的にあとは変わってませんので・・・おいおいこちらにも追記してきます。

あ、あと xlsx に対応させる場合は、lib フォルダの init.vbs 内の記述を以下のように変更してください。

	Const XLS_EXT = ".xlsx"





