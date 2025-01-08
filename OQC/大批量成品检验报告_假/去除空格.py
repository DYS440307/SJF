import os
from openpyxl import load_workbook


def replace_spaces_with_underscores(directory):
    # 遍历目标目录下的所有文件
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        # 只处理.xlsx和.xlsm文件
        if filename.endswith('.xlsx') or filename.endswith('.xlsm'):
            print(f"Processing file: {file_path}")
            # 打开工作簿
            wb = load_workbook(file_path)

            # 修改工作簿的文件名
            new_filename = filename.replace(" ", "_")
            new_file_path = os.path.join(directory, new_filename)

            # 如果文件名有变化，则重命名文件
            if new_filename != filename:
                os.rename(file_path, new_file_path)
                print(f"Renamed file: {filename} -> {new_filename}")


# 目标目录
directory = r"F:\system\Desktop\InPut\NEW"

# 执行替换操作
replace_spaces_with_underscores(directory)
