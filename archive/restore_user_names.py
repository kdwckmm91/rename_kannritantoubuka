import pandas as pd
from pathlib import Path

# 逆置換ルール（置換済みデータを元に戻す）
reverse_replacement_rules = {
    # 無デザ部門の逆置換
    '後藤　健太': [
        ('無デザ', '阿部　友希'),
        ('無デザ', '小林　満'),
        ('無デザ', '中田　涼太'),
    ],
    # RANプロダクト部門の逆置換
    '川上　純平': [
        ('RANプロダクト', '谷村　誠'),
        ('RANプロダクト', '西岡　裕輝'),
    ],
    # RANインテグ部門の逆置換（多数）
    '塚本　裕太': [
        ('RANインテグ', '小田島　祥太'),
        ('RANインテグ', '加藤　遼介'),
        ('RANインテグ', '平井　尊教'),
        ('RANインテグ', '内山　忠'),
        ('RANインテグ', '近松　慎伍'),
        ('RANインテグ', '藤野　伸児'),
        ('RANインテグ', '三原　寛高'),
        ('RANインテグ', '増田　昌史'),
        ('RANインテグ', '井原　泰介'),
        ('RANインテグ', '田中　晋也'),
        ('RANインテグ', '相良　光毅'),
        ('RANインテグ', '相原　聖'),
        ('RANインテグ', '吉村　匠二'),
        ('RANインテグ', '高瀬　健太'),
        ('RANインテグ', '平本　義貴'),
        ('RANインテグ', '竹内　悠'),
        ('RANインテグ', '橋本　謙一'),
        ('RANインテグ', '斎藤　健太郎'),
        ('RANインテグ', '高原　将紀'),
        ('RANインテグ', '中村　徹'),
        ('RANインテグ', '松谷　龍一'),
        ('RANインテグ', '相野　宏文'),
        ('RANインテグ', '坂井　信行'),
        ('RANインテグ', '大場　弘和'),
        ('RANインテグ', '小西　顕也'),
        ('RANインテグ', '大野　禎久'),
        ('RANインテグ', '古沢　祐之'),
        ('RANインテグ', '鬼頭　理香'),
        ('RANインテグ', '宮崎　寛之'),
        ('RANインテグ', '山崎　肇'),
        ('RANインテグ', '山田　太郎'),
        ('RANインテグ', '岡本　誠一郎'),
        ('RANインテグ', '安藤　英浩'),
        ('RANインテグ', '長嶋　嶺'),
        ('RANインテグ', '森岡　康史'),
        ('RANインテグ', '北澤　伸一郎'),
        ('RANインテグ', '宮野　勇人'),
        ('RANインテグ', '住岡　直美'),
        ('RANインテグ', '彦坂　諭志'),
        ('RANインテグ', '須貝　宜嗣'),
        ('RANインテグ', '田中　文子'),
        ('RANインテグ', '山野　滉之'),
        ('RANインテグ', '島野　大雅'),
        ('RAンインテグ', '磯辺　浩子'),
    ]
}

base = Path(r'C:\Temp\ツール\sisantantou\source')
files_to_restore = [
    base / '無デザ_無シ技_台帳_20260618.xlsx',
    base / 'RAN技_無シ技_台帳_20260618.xlsx',
]

print("=" * 80)
print("利用者名の置換を取り消し（元に戻す）")
print("=" * 80)

for file in files_to_restore:
    if not file.exists():
        print(f"⚠ ファイルが見つかりません: {file.name}")
        continue
    
    print(f"\n処理中: {file.name}")
    df = pd.read_excel(file, sheet_name=0)
    
    restore_count = 0
    for current_name, replacements in reverse_replacement_rules.items():
        for dept, original_name in replacements:
            matches = df[df['利用者名'] == current_name]
            if len(matches) > 0:
                count = len(matches)
                df.loc[df['利用者名'] == current_name, '利用者名'] = original_name
                restore_count += count
                print(f"  {current_name} → {original_name}: {count}件")
    
    # 同じファイルパスで上書き
    df.to_excel(file, index=False, sheet_name='Sheet1')
    print(f"✓ {restore_count}件を復元して保存")

print("\n" + "=" * 80)
print("✓ 利用者名の復元完了")
print("=" * 80)
