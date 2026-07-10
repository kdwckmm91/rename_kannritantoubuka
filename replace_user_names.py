import pandas as pd
from pathlib import Path

BASE_DIR = Path(r'C:\Temp\ツール\sisantantou')
RULES_FILE = BASE_DIR / 'replacement_rules.xlsx'

# --- 置換ルールを Excel から読み込む ---
if not RULES_FILE.exists():
    raise FileNotFoundError(
        f'置換ルールファイルが見つかりません: {RULES_FILE}\n'
        'generate_replacement_rules.py を先に実行してください。'
    )

rules_df = pd.read_excel(RULES_FILE, sheet_name='置換ルール')

# 除外フラグが立っている行はスキップ（空欄 = 置換対象、1 = 対象外）
active_rules = rules_df[
    rules_df['除外フラグ'].isna() | (rules_df['除外フラグ'].astype(str).str.strip() == '')
].copy()

excluded = len(rules_df) - len(active_rules)
print(f'置換ルール読み込み: {len(rules_df)} 件中 {len(active_rules)} 件が有効（{excluded} 件除外）')

# 辞書に変換: {置換前利用者名: (部門, 置換後利用者名)}
replacement_rules = {
    str(row['置換前利用者名']).strip(): (
        str(row['部門']).strip(),
        str(row['置換後利用者名']).strip(),
    )
    for _, row in active_rules.iterrows()
}

base = BASE_DIR / 'source'
file_pairs = [
    (
        base / '無デザ_無シ技_台帳_20260618.xlsx',
        base.parent / '無デザ_無シ技_台帳_20260618_name_replaced.xlsx',
    ),
    (
        base / 'RAN技_無シ技_台帳_20260618.xlsx',
        base.parent / 'RAN技_無シ技_台帳_20260618_name_replaced.xlsx',
    ),
]

print('=' * 80)
print('利用者名の置換処理開始')
print('=' * 80)

for source_file, output_file in file_pairs:
    if not source_file.exists():
        print(f'⚠ ファイルが見つかりません: {source_file.name}')
        continue

    if source_file.resolve() == output_file.resolve():
        raise ValueError(f'出力先が source と同一です: {source_file}')

    print(f'\n処理中: {source_file.name}')
    df = pd.read_excel(source_file, sheet_name=0)

    replace_count = 0
    for old_name, (dept, new_name) in replacement_rules.items():
        matches = df[df['利用者名'] == old_name]
        if len(matches) > 0:
            count = len(matches)
            df.loc[df['利用者名'] == old_name, '利用者名'] = new_name
            replace_count += count
            print(f'  {old_name} → {new_name}: {count}件')

    # source は変更せず、新しいファイルとして保存
    df.to_excel(output_file, index=False, sheet_name='Sheet1')
    print(f'✓ {replace_count}件を置換して保存: {output_file.name}')

print('\n' + '=' * 80)
print('✓ 利用者名の置換完了')
print('=' * 80)
