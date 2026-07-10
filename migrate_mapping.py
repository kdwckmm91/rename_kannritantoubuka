"""
mapping_master_updated.xlsx に「突合キー種別」「除外フラグ」列を追加するマイグレーションスクリプト

突合キー種別:
  利用者名          … 台帳の「利用者名」列で突合
  組織管理区分１名  … 台帳の「組織管理区分１名」列で突合
  組織管理区分２名  … 台帳の「組織管理区分２名」列で突合（手動設定用）

除外フラグ:
  空欄 … 通常通り置換対象
  1   … 既にいない / 異動予定のため置換対象外
"""

import pandas as pd
from pathlib import Path

MAPPING_FILE = Path(r"C:\Temp\ツール\sisantantou\mapping_master_updated.xlsx")

# 既知の組織管理区分１名カテゴリ（台帳に登場する試験系カテゴリ）
ORG1_CATEGORIES = {
    "SV", "対向系", "無特", "商用FR", "自前", "負荷",
    "3G_CP", "LTE_CP", "OAM", "伝送路", "L1L2", "GEN",
    "IP-NW", "RAN品", "ツール", "OREC", "衛星",
}


def determine_type(key: str) -> str:
    return "組織管理区分１名" if str(key).strip() in ORG1_CATEGORIES else "利用者名"


df = pd.read_excel(MAPPING_FILE, sheet_name="対応関係")

# すでに移行済みなら何もしない
if "突合キー種別" in df.columns:
    print("✓ すでに移行済みです")
else:
    # 列名変更
    df = df.rename(columns={"キー（利用者名または組織管理区分１名）": "キー"})

    # 突合キー種別を自動判定して2列目に挿入
    df.insert(1, "突合キー種別", df["キー"].apply(determine_type))

    # 除外フラグを末尾に追加（空欄 = 置換対象、1 = 対象外）
    df["除外フラグ"] = None

    df.to_excel(MAPPING_FILE, index=False, sheet_name="対応関係", engine="openpyxl")
    print(f"✓ {MAPPING_FILE.name} を更新しました")

# 確認
df2 = pd.read_excel(MAPPING_FILE, sheet_name="対応関係")
n_user = (df2["突合キー種別"] == "利用者名").sum()
n_org1 = (df2["突合キー種別"] == "組織管理区分１名").sum()
print(f"  列: {df2.columns.tolist()}")
print(f"  利用者名:          {n_user} 件")
print(f"  組織管理区分１名:  {n_org1} 件")
print(f"  行数合計:          {len(df2)}")
print()
print(df2.head(15).to_string())
