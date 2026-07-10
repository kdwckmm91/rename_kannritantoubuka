import pandas as pd
from pathlib import Path

BASE_DIR = Path(r'C:\Temp\ツール\sisantantou')
SOURCE_DIR = BASE_DIR / 'source'

file = SOURCE_DIR / '無デザ_無シ技_台帳_20260618.xlsx'

df = pd.read_excel(file)

# 衛星電話を検索
sat_matches = df[df['組織管理区分１名'] == '衛星']
print(f"衛星電話レコード: {len(sat_matches)}件\n")

# すべての衛星電話レコードを表示
cols = ['利用者名', '組織管理区分１名', '物品名', '設置場所', '管理担当部課名']
print(sat_matches[cols].to_string())
