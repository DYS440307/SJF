import openpyxl

# 读取源文件
source_path = r'E:\System\pic\1.xlsx'
wb_source = openpyxl.load_workbook(source_path)
ws_source = wb_source.active

# 创建/打开目标文件
fo_extract_path = r'E:\System\pic\FO提取.xlsx'
try:
    wb_fo = openpyxl.load_workbook(fo_extract_path)
    ws_fo = wb_fo.active
except FileNotFoundError:
    wb_fo = openpyxl.Workbook()
    ws_fo = wb_fo.active

# 获取源文件的最大列数
max_col = ws_source.max_column
max_row = ws_source.max_row

# 用来记录FO提取文件里的写入位置
write_row = 1
write_col = 1
count_in_col = 0

# 从第二列开始处理
for col in range(2, max_col + 1):
    max_value = None
    max_row_index = None

    # 查找这一列的最大值
    for row in range(1, max_row + 1):
        cell_value = ws_source.cell(row=row, column=col).value
        if isinstance(cell_value, (int, float)):  # 确保是数字
            if (max_value is None) or (cell_value > max_value):
                max_value = cell_value
                max_row_index = row

    if max_row_index is not None:
        # 拿到第一列（A列）对应的值
        corresponding_value = ws_source.cell(row=max_row_index, column=1).value

        # 写入FO提取工作簿
        ws_fo.cell(row=write_row, column=write_col, value=corresponding_value)

        write_row += 1
        count_in_col += 1

        # 如果一列已经写了72个数值，就换到下一列
        if count_in_col >= 72:
            write_col += 1
            write_row = 1
            count_in_col = 0

# 保存结果
wb_fo.save(fo_extract_path)

print("处理完成！")
