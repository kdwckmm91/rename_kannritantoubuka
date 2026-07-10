import pandas as pd
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BASE_DIR = Path(r"C:\Temp\ツール\sisantantou")
SOURCE_DIR = BASE_DIR / "source"
MAPPING_FILE = BASE_DIR / "mapping_master_updated.xlsx"

ledger_files = [
    SOURCE_DIR / "無デザ_無シ技_台帳_20260618.xlsx",
    SOURCE_DIR / "RAN技_無シ技_台帳_20260618.xlsx"
]

logger.info("=" * 80)
logger.info("複数台帳の一括更新開始（組織管理区分１名優先・除外フラグ対応）")
logger.info("=" * 80)

# 対応関係を読み込み
logger.info(f"対応関係ファイルを読み込み: {MAPPING_FILE.name}")
mapping_df = pd.read_excel(MAPPING_FILE, sheet_name="対応関係")
logger.info(f"✓ {len(mapping_df)} 行の対応関係を読み込みました")

# 突合種別ごとの辞書を構築（除外フラグが立っている行はスキップ）
user_dict  = {}   # 利用者名 → (code, name)
org1_dict  = {}   # 組織管理区分１名 → (code, name)
org2_dict  = {}   # 組織管理区分２名 → (code, name)
excluded_count = 0

for _, row in mapping_df.iterrows():
    # 除外フラグが設定されている行はスキップ
    if pd.notna(row.get("除外フラグ")) and str(row["除外フラグ"]).strip() not in ("", "nan"):
        excluded_count += 1
        logger.info(f"  除外: {row['キー']} ({row.get('突合キー種別', '?')})")
        continue

    key       = str(row["キー"]).strip()
    code      = int(row["新しい管理担当部課コード"])
    name      = row["新しい管理担当部課名"]
    key_type  = str(row.get("突合キー種別", "利用者名")).strip()

    if key_type == "組織管理区分１名":
        org1_dict[key] = (code, name)
    elif key_type == "組織管理区分２名":
        org2_dict[key] = (code, name)
    else:  # 利用者名
        user_dict[key] = (code, name)

logger.info(f"✓ 利用者名マッピング:          {len(user_dict)} 件")
logger.info(f"✓ 組織管理区分１名マッピング:  {len(org1_dict)} 件")
logger.info(f"✓ 組織管理区分２名マッピング:  {len(org2_dict)} 件")
logger.info(f"✓ 除外（フラグあり）:          {excluded_count} 件")

# 各台帳ファイルを更新
for ledger_file in ledger_files:
    if not ledger_file.exists():
        logger.warning(f"ファイルが見つかりません: {ledger_file.name}")
        continue

    logger.info(f"\n{'='*80}")
    logger.info(f"台帳ファイルを処理: {ledger_file.name}")
    logger.info(f"{'='*80}")

    ledger_df = pd.read_excel(ledger_file, sheet_name=0)
    logger.info(f"✓ {len(ledger_df):,} 行の台帳を読み込みました")

    updated_org1  = 0
    updated_org2  = 0
    updated_user  = 0
    no_match      = 0

    # 組織管理区分２名列の有無を確認
    has_org2_col = "組織管理区分２名" in ledger_df.columns

    for idx, row in ledger_df.iterrows():
        user_val  = str(row["利用者名"]).strip()          if pd.notna(row.get("利用者名"))          else None
        org1_val  = str(row["組織管理区分１名"]).strip()  if pd.notna(row.get("組織管理区分１名"))  else None
        org2_val  = str(row["組織管理区分２名"]).strip()  if has_org2_col and pd.notna(row.get("組織管理区分２名")) else None

        matched = False

        # 優先1: 組織管理区分１名（試験系機器の指標）
        if org1_val and org1_val in org1_dict:
            code, name = org1_dict[org1_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            updated_org1 += 1
            matched = True

        # 優先2: 組織管理区分２名
        elif org2_val and org2_val in org2_dict:
            code, name = org2_dict[org2_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            updated_org2 += 1
            matched = True

        # 優先3: 利用者名（除外フラグ付きはマッピング辞書から除外済み）
        elif user_val and user_val in user_dict:
            code, name = user_dict[user_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            updated_user += 1
            matched = True

        if not matched:
            no_match += 1

    output_file = BASE_DIR / ledger_file.name.replace(".xlsx", "_updated.xlsx")
    ledger_df.to_excel(output_file, index=False, sheet_name="Sheet1")

    logger.info(f"✓ 更新完了: {updated_org1 + updated_org2 + updated_user} 行")
    logger.info(f"  - 組織管理区分１名で突合: {updated_org1}")
    logger.info(f"  - 組織管理区分２名で突合: {updated_org2}")
    logger.info(f"  - 利用者名で突合:         {updated_user}")
    logger.info(f"  - マッチなし:             {no_match}")
    logger.info(f"✓ 出力ファイル: {output_file.name}")

logger.info(f"\n{'='*80}")
logger.info("✓ すべての台帳を更新完了")
logger.info(f"{'='*80}")

