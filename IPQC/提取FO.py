import openpyxl

# 打开源Excel文件
source_file = r'C:\Users\SL\Downloads\1.xlsx'
wb_source = openpyxl.load_workbook(source_file)
sheet_source = wb_source.active  # 获取活动工作表

# 获取第14行到第29行的数据（从第二列开始）
data = []
for row in range(14, 30):  # 行号从14到29
    row_data = [sheet_source.cell(row=row, column=col).value for col in range(2, sheet_source.max_column + 1)]
    data.append(row_data)

# 找到最大值及其对应的第一列值
max_value = None
corresponding_value = None
for row in data:
    row_max = max(row)
    if max_value is None or row_max > max_value:
        max_value = row_max
        corresponding_value = sheet_source.cell(row=data.index(row) + 14, column=1).value

# 打开目标Excel文件（FO提取）
target_file = r'FO提取.xlsx'
wb_target = openpyxl.load_workbook(target_file)
sheet_target = wb_target.active  # 获取活动工作表

# 将对应的值写入目标工作簿的第一列
sheet_target.append([corresponding_value])

# 保存目标Excel文件
wb_target.save(target_file)

print(f"最大值: {max_value}")
print(f"对应的第一列值: {corresponding_value}")
