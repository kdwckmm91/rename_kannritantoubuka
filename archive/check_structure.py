#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台帳ファイルの構造を確認するスクリプト
"""

import pandas as pd
from pathlib import Path

excel_file = Path(r"C:\Temp\ツール\sisantantou\台帳_20260616.xlsx")

print("=" * 100)
print("台帳ファイル構造確認")
print("=" * 100)

try:
    # ファイルの存在確認
    if not excel_file.exists():
        print(f"エラー: ファイルが見つかりません: {excel_file}")
        exit(1)
    
    print(f"\n✓ ファイル: {excel_file}")
    print(f"  ファイルサイズ: {excel_file.stat().st_size / 1024 / 1024:.2f} MB\n")
    
    # シート情報
    xl_file = pd.ExcelFile(excel_file)
    print(f"シート名: {xl_file.sheet_names}\n")
    
    # 最初のシートを読み込み
    sheet_name = xl_file.sheet_names[0]
    print(f"{'=' * 100}")
    print(f"シート: {sheet_name}")
    print(f"{'=' * 100}\n")
    
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    print(f"行数: {len(df):,}")
    print(f"列数: {len(df.columns)}\n")
    
    print("列一覧:")
    print("-" * 100)
    for i, col in enumerate(df.columns, 1):
        # 列のデータ型と最初の値を表示
        first_val = df[col].iloc[0] if len(df) > 0 else "N/A"
        dtype = df[col].dtype
        print(f"  {i:3d}. [{col:20s}] {str(dtype):15s} | 最初の値: {str(first_val)[:50]}")
    
    # BQ列の確認
    print(f"\n{'=' * 100}")
    print("BQ列（利用者名）の確認")
    print(f"{'=' * 100}\n")
    
    if 'BQ' in [col for col in df.columns if isinstance(col, str) and col.upper() == 'BQ']:
        print("✓ BQ列が見つかりました")
        print(f"  ユニークな値の数: {df['BQ'].nunique()}")
        print(f"  最初の5行:")
        for idx, val in enumerate(df['BQ'].head(), 1):
            print(f"    {idx}. {val}")
    else:
        # Excel列名で確認
        print("⚠ BQ という列が見つかりません。")
        print("   Excel列名との対応を確認中...")
        # 最後から数えてBQ列に相当する位置を確認
        col_index = 67  # BQ = 26*2 + 15 = 67
        if col_index - 1 < len(df.columns):
            bq_col = df.columns[col_index - 1]
            print(f"   Excel の BQ列は DataFrame では: {bq_col}")
            print(f"   最初の5行: {df[bq_col].head().tolist()}")
    
    # C列と D列の確認
    print(f"\n{'=' * 100}")
    print("C列（管理担当部課）と D列（管理担当部課名）の確認")
    print(f"{'=' * 100}\n")
    
    if len(df.columns) >= 3:
        c_col = df.columns[2]  # C列は3番目
        d_col = df.columns[3]  # D列は4番目
        print(f"C列: {c_col}")
        print(f"  最初の5行: {df[c_col].head().tolist()}")
        print(f"\nD列: {d_col}")
        print(f"  最初の5行: {df[d_col].head().tolist()}")
    
    # 「組織管理区分１名」の確認
    print(f"\n{'=' * 100}")
    print("「組織管理区分１名」列の確認")
    print(f"{'=' * 100}\n")
    
    org_col = None
    for col in df.columns:
        if isinstance(col, str) and '組織管理区分' in col and '１' in col:
            org_col = col
            break
    
    if org_col:
        print(f"✓ 見つかりました: {org_col}")
        print(f"  ユニークな値の数: {df[org_col].nunique()}")
        print(f"  最初の5行: {df[org_col].head().tolist()}")
    else:
        print("⚠ 「組織管理区分１名」という名前の列が見つかりません。")
        print("   候補となる列:")
        for col in df.columns:
            if isinstance(col, str) and '組織' in col:
                print(f"     - {col}")
    
    # サンプルデータ表示
    print(f"\n{'=' * 100}")
    print("サンプルデータ（最初の3行）")
    print(f"{'=' * 100}\n")
    print(df.head(3).to_string())
    
except Exception as e:
    print(f"エラーが発生しました: {e}")
    import traceback
    traceback.print_exc()
