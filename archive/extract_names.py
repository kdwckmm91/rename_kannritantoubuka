import pandas as pd
from pathlib import Path

base = Path(r"C:\Temp\ツール\sisantantou")
files = sorted(base.glob("*20260618.xlsx"))

for file in files:
    if file.name not in {"無デザ_無シ技_台帳_20260618.xlsx", "RAN技_無シ技_台帳_20260618.xlsx"}:
        continue

    print("=" * 80)
    print(file.name)
    print("=" * 80)

    df = pd.read_excel(file, sheet_name=0)
    for col_name in ["利用者名", "組織管理区分１名"]:
        if col_name not in df.columns:
            print(f"[{col_name}] 列が見つかりません")
            continue

        values = (
            df[col_name]
            .dropna()
            .astype(str)
            .map(str.strip)
        )
        values = [v for v in values if v and v.lower() != "nan"]
        unique_values = list(dict.fromkeys(values))

        print(f"[{col_name}] {len(unique_values)}件")
        for value in unique_values:
            print(value)
        print()
