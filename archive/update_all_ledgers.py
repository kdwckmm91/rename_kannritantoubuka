#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2つの台帳を対応関係ファイルで一括更新
突合キー種別列（利用者名 / 組織管理区分１名 / 組織管理区分２名）を使って
台帳の適切な列のみを突合対象にする。
除外フラグが立っているエントリは突合対象外とする。
"""

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
logger.info("複数台帳の一括更新開始（突合キー種別・除外フラグ対応）")
logger.info("=" * 80)

# 対応関係を読み込み
logger.info(f"対応関係ファイルを読み込み: {MAPPING_FILE.name}")
mapping_df = pd.read_excel(MAPPING_FILE, sheet_name="対応関係")
logger.info(f"✓ {len(mapping_df)} 行の対応関係を読み込みました")

# 突合種別ごとに辞書を構築（除外フラグ行はスキップ）
user_dict = {}   # 利用者名 → (code, name)
org1_dict = {}   # 組織管理区分１名 → (code, name)
org2_dict = {}   # 組織管理区分２名 → (code, name)
excluded_count = 0

for _, row in mapping_df.iterrows():
    if pd.notna(row.get("除外フラグ")) and str(row["除外フラグ"]).strip() not in ("", "nan"):
        excluded_count += 1
        continue

    key      = str(row["キー"]).strip()
    code     = int(row["新しい管理担当部課コード"])
    name     = row["新しい管理担当部課名"]
    key_type = str(row.get("突合キー種別", "利用者名")).strip()

    if key_type == "組織管理区分１名":
        org1_dict[key] = (code, name)
    elif key_type == "組織管理区分２名":
        org2_dict[key] = (code, name)
    else:
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

    updated_user = 0
    updated_org1 = 0
    updated_org2 = 0
    no_match     = 0

    has_org2_col = "組織管理区分２名" in ledger_df.columns

    for idx, row in ledger_df.iterrows():
        user_val = str(row["利用者名"]).strip()         if pd.notna(row.get("利用者名"))         else None
        org1_val = str(row["組織管理区分１名"]).strip() if pd.notna(row.get("組織管理区分１名")) else None
        org2_val = str(row["組織管理区分２名"]).strip() if has_org2_col and pd.notna(row.get("組織管理区分２名")) else None

        matched = False

        if org1_val and org1_val in org1_dict:
            code, name = org1_dict[org1_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            updated_org1 += 1
            matched = True
        elif org2_val and org2_val in org2_dict:
            code, name = org2_dict[org2_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            updated_org2 += 1
            matched = True
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


logger.info("=" * 80)
logger.info("複数台帳の一括更新開始")
logger.info("=" * 80)

# 対応関係を読み込み
logger.info(f"対応関係ファイルを読み込み: {MAPPING_FILE.name}")
mapping_df = pd.read_excel(MAPPING_FILE, sheet_name='対応関係')
logger.info(f"✓ {len(mapping_df)} 行の対応関係を読み込みました")

# 対応関係を辞書化
mapping_dict = {}
for _, row in mapping_df.iterrows():
    key = str(row['キー（利用者名または組織管理区分１名）']).strip()
    new_code = int(row['新しい管理担当部課コード'])
    new_name = row['新しい管理担当部課名']
    mapping_dict[key] = (new_code, new_name)

logger.info(f"対応関係辞書を作成: {len(mapping_dict)} 件")

# 各台帳ファイルを更新
for ledger_file in ledger_files:
    if not ledger_file.exists():
        logger.warning(f"ファイルが見つかりません: {ledger_file.name}")
        continue

    logger.info(f"\n{'='*80}")
    logger.info(f"台帳ファイルを処理: {ledger_file.name}")
    logger.info(f"{'='*80}")

    # 台帳を読み込み
    ledger_df = pd.read_excel(ledger_file, sheet_name=0)
    logger.info(f"✓ {len(ledger_df):,} 行の台帳を読み込みました")

    updated_count = 0
    matched_by_user = 0
    matched_by_org = 0

    # 各行を処理
    for idx, row in ledger_df.iterrows():
        user_name = str(row['利用者名']).strip() if pd.notna(row['利用者名']) else None
        org_name = str(row['組織管理区分１名']).strip() if pd.notna(row['組織管理区分１名']) else None

        # 利用者名で突合
        if user_name and user_name in mapping_dict:
            new_code, new_name = mapping_dict[user_name]
            ledger_df.at[idx, '管理担当部課'] = new_code
            ledger_df.at[idx, '管理担当部課名'] = new_name
            updated_count += 1
            matched_by_user += 1

        # 組織管理区分１名で突合（ユーザー名で見つからない場合）
        elif org_name and org_name in mapping_dict:
            new_code, new_name = mapping_dict[org_name]
            ledger_df.at[idx, '管理担当部課'] = new_code
            ledger_df.at[idx, '管理担当部課名'] = new_name
            updated_count += 1
            matched_by_org += 1

    # 出力ファイルを生成
    output_file = ledger_file.parent / ledger_file.name.replace('.xlsx', '_updated.xlsx')
    ledger_df.to_excel(output_file, index=False, sheet_name='Sheet1')

    logger.info(f"✓ 更新完了: {updated_count} 行")
    logger.info(f"  - 利用者名で突合: {matched_by_user}")
    logger.info(f"  - 組織管理区分１名で突合: {matched_by_org}")
    logger.info(f"✓ 出力ファイル: {output_file.name}")

logger.info(f"\n{'='*80}")
logger.info("✓ すべての台帳を更新完了")
logger.info(f"{'='*80}")
