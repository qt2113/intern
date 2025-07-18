# -*- coding: utf-8 -*-
"""
rule1.py
定义邮件内容解析规则，返回表格 <td> 数据及表头所在位置（行数）
"""

from bs4 import BeautifulSoup

def chooseMethed(soup: BeautifulSoup, heading: str):
    """
    根据邮件标题选择解析方式。
    返回：
        data: list of <td> elements
        cishu: 表头所在行数或标志
    """
    # 默认处理方式
    table = soup.find('table')
    if not table:
        return None, 0

    # 提取所有td
    tds = table.find_all('td')

    # 根据标题灵活匹配不同格式（可扩展）
    if "产品净值" in heading:
        # 特殊处理产品净值格式（例）
        return tds, 1  # 表头为第一行
    elif "每日净值" in heading:
        return tds, 0  # 表头为第0行
    elif "估值表" in heading:
        return tds, 2  # 表头为第2行
    else:
        return tds, 1  # 默认情况

# def choose_method(raw_data, mapping):
#     """
#     raw_data: DataFrame，包含首行为表头的数据
#     mapping: dict，标准字段到可能表头名称的映射

#     返回值:
#         data_rows: 表格中除去表头的数据行
#         column_map: 标准字段 -> raw_data中对应列的下标
#     """
#     # 取第一行为表头
#     header_row = list(raw_data.iloc[0])
#     data_rows = raw_data.iloc[1:].values.tolist()

#     # 建立标准字段到列号的映射
#     column_map = {}
#     for idx, col_name in enumerate(header_row):
#         if not isinstance(col_name, str):
#             continue
#         col_name = col_name.strip()
#         for standard_field, possible_names in mapping.items():
#             if col_name in possible_names:
#                 column_map[standard_field] = idx
#                 break  # 避免一个表头名匹配到多个字段

#     return data_rows, column_map
