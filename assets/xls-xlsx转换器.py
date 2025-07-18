import os
import win32com.client

def convert_xls_to_xlsx(folder_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    for filename in os.listdir(folder_path):
        if filename.endswith(".xls") and not filename.startswith("~$"):
            full_path = os.path.join(folder_path, filename)
            new_path = os.path.join(folder_path, filename.replace(".xls", ".xlsx"))

            try:
                wb = excel.Workbooks.Open(full_path)
                wb.SaveAs(new_path, FileFormat=51)  # 51 is for xlsx
                wb.Close(False)
                print(f"转换成功: {filename} → {os.path.basename(new_path)}")
            except Exception as e:
                print(f"转换失败: {filename}，错误信息: {e}")

    excel.Quit()

if __name__ == "__main__":
    folder = r"C:\Users\TQY\Desktop\遂玖日出定向2号" # 改成你的文件夹路径
    convert_xls_to_xlsx(folder)
