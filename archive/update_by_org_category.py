import pandas as pd
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BASE_DIR = Path(r"C:\Temp\ツール\sisantantou")
SOURCE_DIR = BASE_DIR / "source"

# 新しいマッピングルール：組織管理区分１に基づく部門移管
org_category_mappings = {
    'RAN品': ('822108000171', '無線ネットワークデザイン部門'),
    '無特': ('822108020008', 'RANプロダクト推進部門'),
}

# 特殊ルール：増田昌史さんの衛星電話は事業企画に移管
special_mappings = {
    ('衛星', '増田　昌史'): ('822108000004', '事業企画部')
}

file_pairs = [
    (
        SOURCE_DIR / "無デザ_無シ技_台帳_20260618.xlsx",
        BASE_DIR / "無デザ_無シ技_台帳_20260618_updated.xlsx"
    ),
    (
        SOURCE_DIR / "RAN技_無シ技_台帳_20260618.xlsx",
        BASE_DIR / "RAN技_無シ技_台帳_20260618_updated.xlsx"
    )
]

logger.info("=" * 80)
logger.info("組織管理区分１に基づく部門移管処理開始（修正版）")
logger.info("=" * 80)

for source_file, output_file in file_pairs:
    if not source_file.exists():
        logger.warning(f"ソースファイルが見つかりません: {source_file.name}")
        continue

    if source_file.resolve() == output_file.resolve():
        raise ValueError(f"出力先が source と同一です: {source_file}")

    logger.info(f"\n処理中: {source_file.name}")
    
    # ファイルを読み込み
    df = pd.read_excel(source_file, sheet_name=0)
    logger.info(f"  ✓ {len(df):,} 行を読み込みました")

    # 管理担当部課列を文字列に変換
    if '管理担当部課' in df.columns:
        df['管理担当部課'] = df['管理担当部課'].astype(str)
    
    update_count_org = 0
    update_count_special = 0
    
    # ルール1: 組織管理区分１に基づく部門移管
    for org_category, (dept_code, dept_name) in org_category_mappings.items():
        # 対象行をフィルタリング
        if '組織管理区分１名' in df.columns:
            mask = df['組織管理区分１名'] == org_category
            count = mask.sum()
            
            if count > 0:
                df.loc[mask, '管理担当部課'] = dept_code
                df.loc[mask, '管理担当部課名'] = dept_name
                update_count_org += count
                logger.info(f"    {org_category} → {dept_name}: {count}件")
    
    # ルール2: 特殊ルール（増田昌史さんの衛星電話）
    for (org_cat, user_name), (dept_code, dept_name) in special_mappings.items():
        if '組織管理区分１名' in df.columns and '利用者名' in df.columns:
            mask = (df['組織管理区分１名'] == org_cat) & (df['利用者名'] == user_name)
            count = mask.sum()
            
            if count > 0:
                df.loc[mask, '管理担当部課'] = dept_code
                df.loc[mask, '管理担当部課名'] = dept_name
                update_count_special += count
                logger.info(f"    {user_name}さんの{org_cat} → {dept_name}: {count}件")

    # source は変更せず、新しいファイルとして保存
    df.to_excel(output_file, index=False, sheet_name='Sheet1')
    
    logger.info(f"\n  ✓ 更新完了:")
    logger.info(f"    - 組織管理区分１ルール: {update_count_org}件")
    logger.info(f"    - 特殊ルール: {update_count_special}件")
    logger.info(f"    - 出力ファイル: {output_file.name}")

logger.info(f"\n{'='*80}")
logger.info("✓ 部門移管処理完了")
logger.info(f"{'='*80}")
