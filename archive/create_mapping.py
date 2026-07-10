#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
対応関係ファイル（mapping_master.xlsx）を作成するスクリプト
"""

import pandas as pd
from pathlib import Path

# ファイルパス
mapping_file = Path(r"C:\Temp\ツール\sisantantou\mapping_master.xlsx")

# 対応関係データの作成（ひな形）
# パターン1: 利用者名をキーとした対応
mapping_data = {
    'キー（利用者名または組織管理区分１名）': [
        '利用者A',
        '利用者B',
        '組織管理区分X',
        '例',
    ],
    '新しい管理担当部課コード': [
        '822108020001',
        '822108020002',
        '822108020003',
        '000000000000',
    ],
    '新しい管理担当部課名': [
        '営業部',
        '企画部',
        '技術部',
        '部課名',
    ]
}

df = pd.DataFrame(mapping_data)

# ひな形ファイルを作成
try:
    df.to_excel(mapping_file, index=False, sheet_name='対応関係')
    print(f"✓ 対応関係ファイルを作成しました: {mapping_file}")
    print(f"\n作成されたファイルの内容:")
    print(df.to_string(index=False))
    print(f"\n使用方法:")
    print("1. 『キー（利用者名または組織管理区分１名）』列に、台帳の利用者名または組織管理区分１名を入力")
    print("2. 『新しい管理担当部課コード』列に、新しい部課コード（数値）を入力")
    print("3. 『新しい管理担当部課名』列に、新しい部課名を入力")
    print("4. main.py を実行して台帳を更新")
except Exception as e:
    print(f"エラー: {e}")
    import traceback
    traceback.print_exc()
