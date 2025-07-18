import pandas as pd

# === 设置文件路径和sheet名（你需要根据你的文件实际路径和sheet名修改） ===
excel_path = r"C:\Users\TQY\Desktop\每股股东类型占比结果.xlsx"
sheet_old = '2024.06.30'
sheet_new = '2025.07.08'

# === 读取两个时间点的数据 ===
df_old = pd.read_excel(excel_path, sheet_name=sheet_old)
df_new = pd.read_excel(excel_path, sheet_name=sheet_new)

# === 提取共同证券代码 ===
common_codes = set(df_old['证券代码']).intersection(set(df_new['证券代码']))

# === 筛选出共同公司 ===
df_old_common = df_old[df_old['证券代码'].isin(common_codes)].copy()
df_new_common = df_new[df_new['证券代码'].isin(common_codes)].copy()

# === 合并两个数据表 ===
df_merged = pd.merge(
    df_old_common, df_new_common,
    on='证券代码',
    suffixes=('_20240630', '_20250708')
)

# === 计算股东占比变化 ===
for col in ['机构占比', '产品占比', '自然人占比', '前十大股东总占比']:
    df_merged[f'{col}变化量'] = df_merged[f'{col}_20250708'] - df_merged[f'{col}_20240630']

# === 选择输出列（可根据需要扩展或删减）===
columns_to_output = [
    '证券代码',
    '证券简称_20240630', 
    '流通市值_20240630', '流通市值_20250708',
    '机构占比_20240630', '机构占比_20250708', '机构占比变化量',
    '产品占比_20240630', '产品占比_20250708', '产品占比变化量',
    '自然人占比_20240630', '自然人占比_20250708', '自然人占比变化量',
    '前十大股东总占比_20240630', '前十大股东总占比_20250708', '前十大股东总占比变化量'
]

df_result = df_merged[columns_to_output]

# === 保存结果到新sheet ===
with pd.ExcelWriter(excel_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df_result.to_excel(writer, sheet_name='股东占比变化分析', index=False)

print("✅ 股东结构变化分析已完成并写入Excel新sheet！")
