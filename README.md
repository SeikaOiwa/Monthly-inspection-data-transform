# 月例点検データの変換ツール
「月例点検アプリ（PowerApp）」の出力データ（月例点検データ）を解析し、月例点検データおよびフロン機器点検データとして保存する

![image](./雛型データ/flow_imge.png)

## 環境構築
(1) Conda環境設定

`conda create -n geturei python`

`conda activate geturei`

(2) モジュールインストール

`pip install pandas`

`pip install openpyxl`

`pip install pywin32`

`pip install pypdf`

`pip install pyinstaller`

## exeファイル作成

「Exe作成」フォルダに`transform_data.py`を保管し、以下のコマンドを実行

`pyinstaller transform_data.py --onefile --noconsole --name ファイル変換 --icon image.ico`

## インプットファイルの要件

- 月例点検結果表：列名（`機器番号,装置名,設置場所,点検月,点検者,安全衛生委員,室長,点検結果備考,点検結果`）を含むこと

- 点検基準表：列名（`該当機器,点検番号,点検部位,点検内容,点検方法,判定基準`）を含む事、`該当機器`は`機器番号`と一致すること

- `点検番号`は、フロン点検の場合には`freon`の文字を含める事（例：`freon_1、freon_a、、`）、フロン点検以外には`freon`の文字を含ませない事

## 実行方法

- 「ファイル名の登録」/ファイル名情報.xlsxを開き、ファイル名と出力先パスを入力

- Monthly-inspection-data-transform内に「月例点検データ」、「点検基準表」を保存

- ファイル変換.exeを実行

- `出力先パス`に〇〇年度_月例点検データ、〇〇年度.odf_フロン月例点検データ.pdfが生成
