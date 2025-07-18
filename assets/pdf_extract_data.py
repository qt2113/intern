# import pdfplumber
# import pandas as pd
# import os

# # 设置输入文件夹路径
# input_folder = r"C:\Users\TQY\Desktop\确认单"
# output_excel = os.path.join(input_folder, "transaction_confirmations.xlsx")

# # 确保文件夹存在
# if not os.path.exists(input_folder):
#     raise FileNotFoundError(f"文件夹 {input_folder} 不存在")

# # 获取所有 PDF 文件
# pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

# # 创建 Excel 写入器
# with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
#     for pdf_file in pdf_files:
#         pdf_path = os.path.join(input_folder, pdf_file)
#         sheet_name = os.path.splitext(pdf_file)[0][:31]  # 限制 sheet 名长度

#         try:
#             all_tables = []
#             with pdfplumber.open(pdf_path) as pdf:
#                 for page in pdf.pages:
#                     tables = page.extract_tables()
#                     for table in tables:
#                         if table and len(table) > 1:
#                             df = pd.DataFrame(table[1:], columns=table[0])
#                             all_tables.append(df)

#             if all_tables:
#                 combined_df = pd.concat(all_tables, ignore_index=True)
#             else:
#                 combined_df = pd.DataFrame([{"信息": "无有效表格"}])

#         except Exception as e:
#             combined_df = pd.DataFrame([{"错误信息": str(e)}])

#         # 写入到 Excel 的 sheet
#         combined_df.to_excel(writer, sheet_name=sheet_name, index=False)

# print(f"✅ 所有PDF解析完成，结果已保存至:\n{output_excel}")



# 第二版
import os
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# 可选：如果 tesseract 没配置环境变量，设置路径（根据你实际安装位置）
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

input_folder = r"C:\Users\TQY\Desktop\数据.pdf"
output_excel = os.path.join(input_folder, "数据提取.xlsx")

# 工具函数：确保列名唯一，避免 reindex 错误
def make_unique(columns):
    seen = {}
    result = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            result.append(col)
    return result

# 检查路径
if not os.path.exists(input_folder):
    raise FileNotFoundError(f"文件夹不存在：{input_folder}")

# 获取所有 PDF 文件
pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        sheet_name = os.path.splitext(pdf_file)[0][:31]  # 限制 sheet 名长度
        all_content = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_number, page in enumerate(pdf.pages, start=1):
                    # 1.表格提取
                    tables = page.extract_tables()
                    has_valid_table = False

                    if tables:
                        for table in tables:
                            if table and len(table) > 1:
                                try:
                                    columns = make_unique(table[0])
                                    df = pd.DataFrame(table[1:], columns=columns)
                                    all_content.append(df)
                                    has_valid_table = True
                                except Exception as e:
                                    # 某些表格行列数不匹配时跳过该表
                                    all_content.append(pd.DataFrame([{
                                        # "页码": page_number,
                                        "识别类型": "表格",
                                        "错误信息": f"表格解析失败: {str(e)}"
                                    }]))

                    # 2.文本提取（若无表格）
                    if not has_valid_table:
                        text = page.extract_text()
                        if text and len(text.strip()) > 10:
                            df = pd.DataFrame([{
                                # "页码": page_number,
                                "识别类型": "文本",
                                "提取结果": text.strip()
                            }])
                            all_content.append(df)
                        else:
                            # 3.OCR 图像识别
                            try:
                                images = convert_from_path(pdf_path, dpi=300, first_page=page_number, last_page=page_number)
                                ocr_text = pytesseract.image_to_string(images[0], lang='chi_sim')
                                df = pd.DataFrame([{
                                    # "页码": page_number,
                                    "识别类型": "OCR",
                                    "提取结果": ocr_text.strip()
                                }])
                                all_content.append(df)
                            except Exception as ocr_error:
                                df = pd.DataFrame([{
                                    # "页码": page_number,
                                    "识别类型": "OCR",
                                    "错误信息": f"OCR失败: {str(ocr_error)}"
                                }])
                                all_content.append(df)

            if all_content:
                combined_df = pd.concat(all_content, ignore_index=True)
            else:
                combined_df = pd.DataFrame([{"信息": "无可提取数据"}])

        except Exception as e:
            combined_df = pd.DataFrame([{"错误信息": str(e)}])

        # 写入 Excel 的对应 Sheet
        combined_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"所有PDF解析完成，结果保存在：\n{output_excel}")



