# 台帳管理担当部課一括変換ツール

台帳ファイルの 管理担当部課 と 管理担当部課名 を、対応関係ファイルにもとづいて一括更新する Python ツールです。

現行の正式エントリーポイントは main.py です。

## 概要

- 対応関係は mapping_master.xlsx から読み込みます。
- 処理対象は input フォルダ配下の xlsx ファイルです。
- 更新後ファイルは output フォルダに同名で出力されます。
- source フォルダに同名ファイルがある場合のみ、差分ファイルを自動生成します。
- 実行ログには突合件数、非突合件数、差分実行有無を出力します。

## クイックスタート

### 1. 前提環境

- Python 3.8 以上を推奨
- 必要ライブラリ: pandas, openpyxl

インストール例:

```bash
pip install pandas openpyxl
```

### 2. ファイルを配置

- 処理対象台帳を input に配置
- 比較元の原本台帳を source に配置
- mapping_master.xlsx をルートに配置

### 3. 実行

```bash
cd c:\Temp\ツール\sisantantou
python main.py
```

仮想環境を使用する場合:

```bash
cd c:\Temp\ツール\sisantantou
.venv\Scripts\python.exe main.py
```

### 4. 出力確認

- output 配下に更新済み台帳を出力
- source に同名台帳がある場合は、追加で output/台帳名_変更箇所.xlsx を生成
- source に同名台帳がある場合は、output/台帳名.xlsx の変更セルを赤字化

## ディレクトリ構成

```text
sisantantou/
├── main.py
├── mapping_master.xlsx
├── replacement_rules.xlsx
├── input/                      # 処理対象台帳（入力）
├── source/                     # 比較元の原本台帳（読み取り専用）
├── output/                     # main.py の出力先
├── archive/                    # 旧スクリプト保管先
├── create_diff_optimized.py
├── generate_replacement_rules.py
├── replace_user_names.py
├── restore_user_names_from_diff.py
├── migrate_mapping.py
├── ARCHITECTURE.md
└── README.md
```

## mapping_master.xlsx 仕様

### 対応関係シート

必須シート名: 対応関係

必須列:

| 列名 | 内容 | 例 |
| --- | --- | --- |
| キー | 突合値（利用者名または組織管理区分） | 佐藤太郎 / 無特 |
| 突合キー種別 | 利用者名 / 組織管理区分１名 / 組織管理区分２名 | 組織管理区分１名 |
| 新しい管理担当部課コード | 更新先コード（数値） | 822108020008 |
| 新しい管理担当部課名 | 更新先名称 | RANプロダクト推進部門 |
| 除外フラグ | 空欄なら有効、値ありなら除外 | 1 |

補足:

- 部課コードが数値変換できない行は警告を出してスキップします。
- 除外フラグがある行はマッピング対象外です。

### 台帳同名シートによる設置場所突合

mapping_master.xlsx に、台帳ファイル名と同名のシートがある場合、次の列で設置場所突合を行います。

| 列名 |
| --- |
| 設置場所 |
| 設置場所詳細 |
| 新しい管理担当部課コード |
| 新しい管理担当部課名 |

設置場所突合は、設置場所詳細が空でない行のみ対象です。

## 更新ロジック（優先順位）

main.py は各行に対して、次の順で最初に一致したルールを適用します。

1. SPECIAL_RULES（組織管理区分１名 × 利用者名）
2. OVERRIDE_RULES（組織管理区分１名、次に組織管理区分２名）
3. 対応関係シートの 組織管理区分１名
4. 対応関係シートの 組織管理区分２名
5. 対応関係シートの 利用者名
6. 台帳同名シートの 設置場所 × 設置場所詳細

どれにも一致しない場合は変更しません。

## 補助スクリプト

### 置換ルールファイル生成

generate_replacement_rules.py は replacement_rules.xlsx を生成します。

```bash
python generate_replacement_rules.py
```

### 利用者名置換

replace_user_names.py は replacement_rules.xlsx を参照して、source 台帳から利用者名置換済みファイルを生成します。

```bash
python replace_user_names.py
```

注意:

- 現在の実装はファイル名が固定されているため、対象台帳名が異なる場合はスクリプト修正が必要です。
- 出力先はルート直下の name_replaced 付きファイルです。

### 差分ファイルにもとづく利用者名復元

restore_user_names_from_diff.py は、変更箇所ファイルを使って特定利用者名を復元した新規台帳を output に生成します。

```bash
python restore_user_names_from_diff.py
```

注意:

- 対象ファイル名と対象利用者名はスクリプト内定数で固定です。

### 単体差分処理

create_diff_optimized.py は source と output の同名台帳を比較し、変更箇所ファイルを生成します。

```bash
python create_diff_optimized.py
```

## トラブルシューティング

### 対応関係ファイルが見つからない

症状:

- 対応関係ファイルが見つかりません

確認:

- ルートに mapping_master.xlsx が存在するか
- シート名が 対応関係 になっているか

### input に台帳がない

症状:

- input/ に処理対象の .xlsx ファイルがありません

確認:

- input フォルダに xlsx を配置したか
- 一時ファイル（~$ で始まる）だけになっていないか

### 更新されない行がある

確認:

- キーの表記揺れ（前後空白、全角半角）
- 突合キー種別の指定ミス
- 除外フラグの設定有無
- 優先順位の上位ルールに先に一致していないか

### 差分ファイルが生成されない

確認:

- source に output と同名の台帳があるか
- 管理番号 列が source/output 両方に存在するか

## 注意事項

- source は原本扱いです。直接編集しない運用を推奨します。
- main.py は output を上書き更新します。同名出力がある場合は再生成されます。
- 旧スクリプト群は archive に保管されており、通常運用では使用しません。

## 関連ドキュメント

- 実装詳細は ARCHITECTURE.md を参照してください。
- ただし、運用仕様の正本は main.py です。

## 更新履歴

- 2026-07-10: README を現行実装に同期（main.py 中心の運用手順へ再構成、旧マクロ手順と固定ファイル前提を整理、mapping 仕様と優先順位と補助スクリプト手順を追記）
