import pandas as pd
from pathlib import Path

BASE_DIR = Path(r'C:\Temp\ツール\sisantantou')
SOURCE_DIR = BASE_DIR / 'source'

files = [
    SOURCE_DIR / '無デザ_無シ技_台帳_20260618.xlsx',
    SOURCE_DIR / 'RAN技_無シ技_台帳_20260618.xlsx'
]

print("=" * 80)
print("衛星電話データの検索")
print("=" * 80)

for file in files:
    print(f"\n{file.name}")
    df = pd.read_excel(file)
    
    # 衛星電話を検索
    if '組織管理区分１名' in df.columns:
        sat_matches = df[df['組織管理区分１名'] == '衛星']
        print(f"  衛星電話レコード: {len(sat_matches)}件")
        if len(sat_matches) > 0:
            print("  利用者名の内訳:")
            print(sat_matches['利用者名'].value_counts().to_string())
            print(f"\n  増田昌史さん: {(sat_matches['利用者名'] == '増田　昌史').sum()}件")
