"""
台帳の管理担当部課を一括更新する統合スクリプト

処理優先順位（高い順）:
  1. SPECIAL_RULES  : (組織管理区分１名, 利用者名) の組合せによる個別移管
  2. OVERRIDE_RULES : 組織管理区分列の値によるハードコード移管
  3. mapping_master.xlsx の 組織管理区分１名 突合
  4. mapping_master.xlsx の 組織管理区分２名 突合
  5. mapping_master.xlsx の 利用者名 突合（除外フラグ行はスキップ）

出力: output/<source と同名ファイル>.xlsx
"""

import pandas as pd
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

BASE_DIR     = Path(r"C:\Temp\ツール\sisantantou")
SOURCE_DIR   = BASE_DIR / "source"
OUTPUT_DIR   = BASE_DIR / "output"
MAPPING_FILE = BASE_DIR / "mapping_master.xlsx"

LEDGER_FILES = [
    SOURCE_DIR / "無デザ_無シ技_台帳_20260618.xlsx",
    SOURCE_DIR / "RAN技_無シ技_台帳_20260618.xlsx",
]

# ---------------------------------------------------------------------------
# ハードコードルール（mapping_master.xlsx より優先）
# ---------------------------------------------------------------------------

# 特殊ルール: (組織管理区分１名, 利用者名) → (部課コード, 部課名)
SPECIAL_RULES: dict[tuple, tuple] = {
    ("衛星", "増田　昌史"): (822108000004, "事業企画部"),
}

# 上書きルール: (対象列名, 値) → (部課コード, 部課名)
OVERRIDE_RULES: list[tuple] = [
    ("組織管理区分１名", "RAN品",  822108000171, "無線ネットワークデザイン部門"),
    ("組織管理区分１名", "無特",   822108020008, "RANプロダクト推進部門"),
    ("組織管理区分２名", "4GvRAN", 822108020004, "RANインテグレーション部門"),
]

OVERRIDE_DICT: dict[tuple, tuple] = {
    (col, val): (code, name)
    for col, val, code, name in OVERRIDE_RULES
}

# ---------------------------------------------------------------------------
# mapping_master.xlsx を読み込んで突合辞書を構築
# ---------------------------------------------------------------------------
logger.info("=" * 80)
logger.info("台帳一括更新開始（統合版）")
logger.info("=" * 80)

logger.info(f"対応関係ファイルを読み込み: {MAPPING_FILE.name}")
mapping_df = pd.read_excel(MAPPING_FILE, sheet_name="対応関係")
logger.info(f"✓ {len(mapping_df)} 行の対応関係を読み込みました")

user_dict: dict[str, tuple]  = {}
org1_dict: dict[str, tuple]  = {}
org2_dict: dict[str, tuple]  = {}
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

logger.info(f"  利用者名マッピング:          {len(user_dict)} 件")
logger.info(f"  組織管理区分１名マッピング:  {len(org1_dict)} 件")
logger.info(f"  組織管理区分２名マッピング:  {len(org2_dict)} 件")
logger.info(f"  除外フラグあり:              {excluded_count} 件")
logger.info(f"  上書きルール（OVERRIDE）:    {len(OVERRIDE_DICT)} 件")
logger.info(f"  特殊ルール（SPECIAL）:       {len(SPECIAL_RULES)} 件")

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 各台帳ファイルを更新
# ---------------------------------------------------------------------------
for ledger_file in LEDGER_FILES:
    if not ledger_file.exists():
        logger.warning(f"ファイルが見つかりません: {ledger_file.name}")
        continue

    logger.info(f"\n{'='*80}")
    logger.info(f"台帳ファイルを処理: {ledger_file.name}")
    logger.info(f"{'='*80}")

    ledger_df = pd.read_excel(ledger_file, sheet_name=0)
    logger.info(f"✓ {len(ledger_df):,} 行を読み込みました")

    has_org2_col = "組織管理区分２名" in ledger_df.columns

    cnt_special  = 0
    cnt_override = 0
    cnt_org1     = 0
    cnt_org2     = 0
    cnt_user     = 0
    cnt_no_match = 0

    for idx, row in ledger_df.iterrows():
        user_val = str(row["利用者名"]).strip()         if pd.notna(row.get("利用者名"))         else None
        org1_val = str(row["組織管理区分１名"]).strip() if pd.notna(row.get("組織管理区分１名")) else None
        org2_val = str(row["組織管理区分２名"]).strip() if has_org2_col and pd.notna(row.get("組織管理区分２名")) else None

        matched = False

        # 優先1: SPECIAL_RULES （組織管理区分１名 × 利用者名 の組合せ）
        if not matched and org1_val and user_val:
            result = SPECIAL_RULES.get((org1_val, user_val))
            if result:
                ledger_df.at[idx, "管理担当部課"]  = result[0]
                ledger_df.at[idx, "管理担当部課名"] = result[1]
                cnt_special += 1
                matched = True

        # 優先2: OVERRIDE_RULES （組織管理区分１名 → 組織管理区分２名 の順に確認）
        if not matched:
            for col_name, val in [("組織管理区分１名", org1_val), ("組織管理区分２名", org2_val)]:
                if val:
                    result = OVERRIDE_DICT.get((col_name, val))
                    if result:
                        ledger_df.at[idx, "管理担当部課"]  = result[0]
                        ledger_df.at[idx, "管理担当部課名"] = result[1]
                        cnt_override += 1
                        matched = True
                        break

        # 優先3: mapping_master 組織管理区分１名
        if not matched and org1_val and org1_val in org1_dict:
            code, name = org1_dict[org1_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            cnt_org1 += 1
            matched = True

        # 優先4: mapping_master 組織管理区分２名
        if not matched and org2_val and org2_val in org2_dict:
            code, name = org2_dict[org2_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            cnt_org2 += 1
            matched = True

        # 優先5: mapping_master 利用者名（除外フラグ付きは辞書から除外済み）
        if not matched and user_val and user_val in user_dict:
            code, name = user_dict[user_val]
            ledger_df.at[idx, "管理担当部課"]  = code
            ledger_df.at[idx, "管理担当部課名"] = name
            cnt_user += 1
            matched = True

        if not matched:
            cnt_no_match += 1

    # source と同じファイル名で output/ に出力
    output_file = OUTPUT_DIR / ledger_file.name
    ledger_df.to_excel(output_file, index=False, sheet_name="Sheet1")

    total = cnt_special + cnt_override + cnt_org1 + cnt_org2 + cnt_user
    logger.info(f"✓ 更新完了: {total:,} 行")
    logger.info(f"  - 特殊ルール（SPECIAL）:     {cnt_special}")
    logger.info(f"  - ハードコード上書き（OVERRIDE）: {cnt_override}")
    logger.info(f"  - 組織管理区分１名突合:      {cnt_org1}")
    logger.info(f"  - 組織管理区分２名突合:      {cnt_org2}")
    logger.info(f"  - 利用者名突合:              {cnt_user}")
    logger.info(f"  - マッチなし:                {cnt_no_match}")
    logger.info(f"✓ 出力ファイル: output/{output_file.name}")

logger.info(f"\n{'='*80}")
logger.info("✓ 台帳一括更新完了")
logger.info(f"{'='*80}")
