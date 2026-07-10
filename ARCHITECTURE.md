# ARCHITECTURE — sisantantou ツール

## ディレクトリ構成

```text
sisantantou/
├── source/                         # 読み取り専用の原本台帳（絶対に上書きしない）
│   ├── 無デザ_無シ技_台帳_20260618.xlsx
│   └── RAN技_無シ技_台帳_20260618.xlsx
├── output/                         # 生成ファイルの出力先（main.py の出力先）
├── mapping_master.xlsx             # main.py 用マッピング（利用者名 / 組織管理区分１名 → 部課コード）
├── mapping_master_updated.xlsx     # 実データ入りマッピング 126 行（update_* 系が参照）
│
├── main.py                         # ★ エントリーポイント（下記参照）
├── update_by_org_category.py       # 組織管理区分１ベースの部門移管（最新の一括更新スクリプト）
├── create_diff_optimized.py        # 差分レビューファイル生成（変更行のみ・赤字）
├── replace_user_names.py           # ★マーク担当者を現担当者に置換（非破壊）
├── restore_user_names.py           # replace_user_names.py の逆操作
│
├── update_all_ledgers.py           # [旧版] 両台帳を mapping_master_updated.xlsx で一括更新
├── update_ledgers_org_priority.py  # [旧版] 組織管理区分１名優先の一括更新
├── create_diff_highlight.py        # [旧版] 差分をセル塗りつぶし（処理重・非推奨）
│
├── extract_names.py                # [セットアップ] 台帳から利用者名・組織区分名を抽出
├── generate_mapping.py             # [セットアップ] 部門データから mapping_master_updated.xlsx を生成
├── create_mapping.py               # [セットアップ] mapping_master.xlsx テンプレート生成
├── check_structure.py              # [調査] 台帳の列構造確認
├── check_satellite.py              # [調査] 衛星電話レコードの一覧確認
├── check_satellite_detail.py       # [調査] 衛星電話レコードの詳細確認
└── create_macro_template.py        # [調査] Excel マクロ用テンプレート生成
```

---

## エントリーポイント: main.py

```text
main.py
  │
  ├─ load_mapping()        mapping_master.xlsx を読み込み、辞書 {キー → (部課コード, 部課名)} を構築
  │                        キー = 「利用者名」または「組織管理区分１名」
  │
  ├─ load_ledger()         source/ にある台帳を読み取り専用で読み込み
  │
  ├─ update_ledger()       行ごとに突合
  │                          1st: 利用者名 でマッピング辞書を検索
  │                          2nd: 組織管理区分１名 でフォールバック検索
  │                        ヒットした行の「管理担当部課」「管理担当部課名」を上書き
  │
  └─ save_ledger()         output/ に新規 Excel として保存（source は変更しない）
```

### 入力

| ファイル | 役割 |
| --- | --- |
| `source/無デザ_無シ技_台帳_20260618.xlsx` | 原本台帳（読み取り専用） |
| `mapping_master.xlsx` | 部課コード対応表 |

### 出力

| ファイル | 役割 |
| --- | --- |
| `output/無デザ_無シ技_台帳_20260618_updated.xlsx` | 更新済み台帳 |

---

## 処理フロー全体（推奨実行順）

```text
[準備]
  extract_names.py          → 台帳から利用者名・組織区分名を洗い出し
  generate_mapping.py       → mapping_master_updated.xlsx を生成

[本処理]
  replace_user_names.py     → ★マーク担当者を現担当者に置換して _name_replaced.xlsx を生成
  ↓
  update_by_org_category.py → 組織管理区分１に基づく部門移管を適用して _updated.xlsx を生成
  （または main.py で mapping_master.xlsx ベースの突合更新）
  ↓
  create_diff_optimized.py  → source と updated を比較し、変更行のみを赤字でまとめた _変更箇所.xlsx を生成
```

---

## スクリプト別 入出力サマリ

| スクリプト | 入力 | 出力 | 備考 |
| --- | --- | --- | --- |
| `main.py` | `source/*.xlsx`, `mapping_master.xlsx` | `output/*_updated.xlsx` | エントリーポイント |
| `update_by_org_category.py` | `source/*.xlsx` | `*_updated.xlsx` | RAN品→無デザ、無特→RANプロダクト移管 |
| `create_diff_optimized.py` | `source/*.xlsx`, `*_updated.xlsx` | `*_変更箇所.xlsx` | 変更行のみ・赤字・軽量版 |
| `replace_user_names.py` | `source/*.xlsx` | `*_name_replaced.xlsx` | ★担当者を置換（非破壊） |
| `restore_user_names.py` | `source/*.xlsx` | `source/*.xlsx` 上書き | replace の逆操作 |
| `extract_names.py` | `source/*.xlsx` or BASE_DIR | stdout | セットアップ時の調査用 |
| `generate_mapping.py` | ハードコードされたデータ | `mapping_master_updated.xlsx` | セットアップ時の一回限り |
| `create_mapping.py` | ハードコードされたデータ | `mapping_master.xlsx` | テンプレート生成 |

---

## 原則

- **source/ は読み取り専用** — すべてのスクリプトは source を変更せず、新規ファイルに出力する
- **output/ が生成物置き場** — main.py の出力先。他スクリプトの出力は BASE_DIR 直下に出す設計も可
- `mapping_master.xlsx` ← `main.py` 用（利用者名/組織区分名 キー）
- `mapping_master_updated.xlsx` ← `update_*` 系が参照する実データ入りマッピング
