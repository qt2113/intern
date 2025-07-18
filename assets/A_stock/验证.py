import pandas as pd
import re

# --- 分类函数 ---
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


# --- 主函数：按股票输出每个股东的分类 ---
def extract_and_classify_with_stock(filepath, output_path):
    df = pd.read_excel(filepath, engine='openpyxl')

    code_col = '证券代码'
    name_col = '证券简称'

    result = []

    for _, row in df.iterrows():
        stock_code = row[code_col]
        stock_name = row[name_col]

        for i in range(1, 11):
            name_cols = [col for col in df.columns if f'排名] 第{i}名' in col and '名称' in col]
            for col in name_cols:
                holder_name = str(row.get(col, '')).strip()
                if holder_name and holder_name != 'nan':
                    category = classify_holder(holder_name)
                    result.append({
                        '证券代码': stock_code,
                        '证券简称': stock_name,
                        '股东名称': holder_name,
                        '分类': category
                    })

    result_df = pd.DataFrame(result)
    result_df.to_excel(output_path, index=False)
    print(f"已提取并分类 {len(result_df)} 个股东记录，结果保存在：{output_path}")

# --- 使用示例 ---
input_path = r"C:\Users\TQY\Desktop\全部A股-20240630-前10名股东.xlsx"  # 替换为你的文件路径
output_path = r'C:\Users\TQY\Desktop\每个股东所属股票及分类结果.xlsx'  # 输出路径

extract_and_classify_with_stock(input_path, output_path)
