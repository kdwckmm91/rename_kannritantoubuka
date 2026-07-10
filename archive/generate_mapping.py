import pandas as pd
from pathlib import Path

# 提供されたデータを構造化
data = {
    '無線ネットワークデザイン部門': {
        'code': 822108000171,
        'users': [
            '阿部　友希', '鵜澤　恭宏', '後藤　健太', '小林　満',
            '内藤　耕太', '中田　涼太', '佐々木　希実', '吉川　侑希'
        ],
        'orgs': []
    },
    'RANプロダクト推進部門': {
        'code': 822108020008,
        'users': [
            '川上　純平', '槇田　智史', '川合　裕之', '伊藤　香貴',
            '谷村　誠', '西岡　裕輝', '石神　美穂'
        ],
        'orgs': []
    },
    'RANインテグレーション部門': {
        'code': 822108020004,
        'users': [
            '小田島　祥太', '加藤　遼介', '菊地　裕大', '平井　尊教',
            '高木　由紀子', '瀬下　拓也', '田村　直樹', '清水　良一',
            '内山　忠', '近松　慎伍', '佐藤　圭', '藤野　伸児',
            '塚本　裕太', '岑　俊成', '白岩　和剛', '三原　寛高',
            '増田　昌史', '井原　泰介', '田中　晋也', '相良　光毅',
            '相原　聖', '吉村　匠二', '高瀬　健太', '平本　義貴',
            '竹内　悠', '橋本　謙一', '河辺　泰宏', '斎藤　健太郎',
            '高原　将紀', '中村　徹', '阿部　拓実', '松谷　龍一',
            '相野　宏文', '坂井　信行', '大場　弘和', '小西　顕也',
            '大野　禎久', '古沢　祐之', '鬼頭　理香', '宮崎　寛之',
            '吉川　裕耀', '山崎　肇', '山田　太郎', '野中　信秀',
            '岡本　誠一郎', '清水　大智', '安藤　英浩', '長嶋　嶺',
            '森岡　康史', '北澤　伸一郎', '宮野　勇人', '住岡　直美',
            '石川　裕太', '彦坂　諭志', '田中　敬史', '須貝　宜嗣',
            '田中　文子', '山野　滉之', '島野　大雅', '磯辺　浩子'
        ],
        'orgs': []
    }
}

# 組織管理区分１名のリスト（どの部門に属するかは未指定なので、すべてに適用）
orgs_common = [
    'SV', '対向系', '無特', '商用FR', '自前', '負荷',
    '3G_CP', 'LTE_CP', 'OAM', '伝送路', 'L1L2', 'GEN',
    'IP-NW', 'RAN品', 'ツール', 'OREC', '衛星'
]

# 対応関係リストを生成
mapping_rows = []

for dept_name, dept_info in data.items():
    code = dept_info['code']
    
    # 利用者名をキーとした行を追加
    for user in dept_info['users']:
        mapping_rows.append({
            'キー（利用者名または組織管理区分１名）': user,
            '新しい管理担当部課コード': code,
            '新しい管理担当部課名': dept_name
        })
    
    # 組織管理区分１名をキーとした行を追加
    for org in orgs_common:
        mapping_rows.append({
            'キー（利用者名または組織管理区分１名）': org,
            '新しい管理担当部課コード': code,
            '新しい管理担当部課名': dept_name
        })

# DataFrameを作成
df = pd.DataFrame(mapping_rows)

# mapping_master.xlsx に上書き保存
output_file = Path(r'C:\Temp\ツール\sisantantou\mapping_master_updated.xlsx')
df.to_excel(output_file, index=False, sheet_name='対応関係')

print(f"✓ 更新されたマッピングファイルを作成: {output_file}")
print(f"  総行数: {len(df)}")
print(f"  利用者名エントリ: {sum(1 for row in mapping_rows if row['キー（利用者名または組織管理区分１名）'] in data['無線ネットワークデザイン部門']['users'] or row['キー（利用者名または組織管理区分１名）'] in data['RANプロダクト推進部門']['users'] or row['キー（利用者名または組織管理区分１名）'] in data['RANインテグレーション部門']['users'])}")
print(f"  組織管理区分１名エントリ: {len(orgs_common) * 3}")
print(f"\nサンプル（最初の10行）:")
print(df.head(10).to_string(index=False))
