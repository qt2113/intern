import pandas as pd
import re
from openpyxl import load_workbook

# --- 股东分类函数 ---
def classify_holder(name):
    if pd.isna(name):
        return '未知'

    name = str(name).strip()
    name_cn = name  # 中文原样
    name_en = name  # 英文检查时处理

    # --- 1. 判断产品 ---
    product_keywords = [
        '基金', '资管', '理财', '专户', '计划', '组合', '产品', '契约型',
        '集合资金信托', '集合资产管理计划', '养老金产品', '保险产品', '封闭式基金',
        '集合计划', '资产支持证券', '信托计划', '投资基金', '创业投资', '私募股权','自有资金','ETF','FUND','L.P.','LP'
    ]
    if any(k in name_cn for k in product_keywords):
        # 防止“基金管理有限公司”被错判成产品
        institution_suffixes = [
            '投资有限公司', '管理有限公司', '有限公司', '资产管理有限公司', '资产管理公司',
            '基金管理公司', '信托公司', '证券公司', '银行', '财务公司', '保险',
            '资本', '创投', '集团', '国资', '基金管理人', '有限合伙',
            '合伙企业', '事务所', '账户', '发展中心'
        ]
        if not any(name_cn.endswith(suffix) for suffix in institution_suffixes):
            return '产品'

    # --- 2. 判断机构（中文） ---
    institution_keywords = [
        '投资有限公司', '管理有限公司', '有限公司', '资产管理有限公司', '资产管理公司',
        '基金管理公司', '信托公司', '证券公司', '银行', '财务公司', '保险',
        '资本', '创投', '集团', '国资', '基金管理人', '有限合伙', '合伙企业','管理中心','株式会社','大学','社','农场','管理处',
        '事务所', '账户', '发展中心','公司','管理站','财政局','财政厅','办公室','研究所','科学院','研究院','中心','委员会','医院','投资局','制药厂'
    ]
    if any(name_cn.endswith(k) or k in name_cn for k in institution_keywords):
        return '机构'

    # --- 3. 判断英文机构 ---
    english_institution_keywords = [
        'BANK', 'SECURITIES', 'TRUST', 'INVESTMENT', 'ASSETMANAGEMENT', 'LIMITED', 'LIMITED.', 'INCORPORATED', 'INCORPORATED.', 'INTERNATIONAL', 'LIMITSD'
    ]
    english_suffixes = [
        'LTD', 'INC', 'LLC', 'CO.,LTD.', 'PLC', 'INC.', 'PLC.', 'CORP.', 'BHD.', 'LTD.','CORPORATION', 'COMPANY', 'BANKING','INCORPORATED', 'SA', 'AG', 'ED'
    ]
    well_known_institutions = [
        'UBS AG', 'MORGAN STANLEY', 'JPMORGAN CHASE BANK N.A.',
        'GOLDMAN SACHS', 'HSBC', 'CREDIT SUISSE', 'CITI','SA','AG'
    ]
    # 统一处理英文名，去除空格和大小写
    name_en_upper = name_en.upper().replace(' ', '')
    # 检查英文机构关键词和后缀（忽略空格和大小写）
    if (any(k in name_en_upper for k in english_institution_keywords)
        or any(name_en_upper.endswith(s.replace(' ', '')) for s in english_suffixes)
        or any(name_en_upper.endswith(w.replace(' ', '')) for w in well_known_institutions)
        or name_en_upper in (w.upper().replace(' ', '') for w in well_known_institutions)):
        return '机构'

    # --- 4. 判断自然人（仅中文 2~4 个汉字） ---
    if re.fullmatch(r'[\u4e00-\u9fa5]{2,4}', name_cn):
        return '自然人'

    # --- 5. 默认兜底 ---
    return '自然人'


# --- 主函数 ---
def analyze_shareholders(filepath):
    df = pd.read_excel(filepath, engine='openpyxl')

    # 基础列识别（自动提取）
    code_col = '证券代码'
    name_col = '证券简称'
    mkt_cap_col = [col for col in df.columns if '流通市值' in col][0]

    result_rows = []

    for _, row in df.iterrows():
        code = row[code_col]
        stock_name = row[name_col]
        try:
            mkt_cap = float(str(row[mkt_cap_col]).replace(',', '').strip())
        except:
            mkt_cap = None

        total_ratio = {'机构': 0.0, '产品': 0.0, '自然人': 0.0}

        for i in range(1, 11):
            try:
                holder_name_col = [col for col in df.columns if f'排名] 第{i}名' in col and '名称' in col][0]
                holder_ratio_col = [col for col in df.columns if f'排名] 第{i}名' in col and '比例' in col][0]
                holder_name = row[holder_name_col]
                ratio = row[holder_ratio_col]

                holder_type = classify_holder(holder_name)
                try:
                    ratio_val = float(str(ratio).replace('%', '').replace(',', '').strip())
                except:
                    ratio_val = 0.0
                if holder_type in total_ratio:
                    total_ratio[holder_type] += ratio_val
            except:
                continue

        total_sum = sum(total_ratio.values())
        result_rows.append({
            '证券代码': code,
            '证券简称': stock_name,
            '流通市值': mkt_cap,
            '机构占比': round(total_ratio['机构'], 4),
            '产品占比': round(total_ratio['产品'], 4),
            '自然人占比': round(total_ratio['自然人'], 4),
            '前十大股东总占比': round(total_sum, 4)
        })

    result_df = pd.DataFrame(result_rows)
    return result_df

# --- 使用示例（输入路径 & 输出路径） ---
input_path = r"C:\Users\TQY\Desktop\全部A股-20240630-前10名股东.xlsx"  # 修改为你的文件路径
output_path = r'C:\Users\TQY\Desktop\每股股东类型占比结果.xlsx'  # 输出结果路径
sheet_name = '股东类型占比'  # 新建的sheet名称

output_df = analyze_shareholders(input_path)


try:
    # 如果文件已存在，追加新sheet
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        output_df.to_excel(writer, index=False, sheet_name=sheet_name)
except FileNotFoundError:
    # 文件不存在则新建
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name=sheet_name)

print(f"统计完成！结果已保存为：{output_path}，Sheet名为：{sheet_name}")
