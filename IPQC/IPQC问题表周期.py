import shutil
import os
import openpyxl

source_file = r"F:\system\Desktop\PY\IPQC\第模板周巡线报告.xlsx"
destination_folder = r"F:\system\Desktop\PY\IPQC"

# 创建目标文件夹（如果不存在）
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# 复制文件到目标文件夹中
for i in range(1, 54):  # 复制53份
    destination_file = os.path.join(destination_folder, f"第{i}周巡线报告.xlsx")
    shutil.copy2(source_file, destination_file)

    # 提取文件名并写入A1单元格
    wb = openpyxl.load_workbook(destination_file)
    sheet = wb.active
    file_name = os.path.basename(destination_file)
    sheet['A1'] = f"第{i}周巡线报告"
    wb.save(destination_file)

print("复制并写入文件名完成！")
