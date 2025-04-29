import openpyxl

# 打开源文件
source_path = r'E:\System\download\4.19批次_KL主音箱扫频后测试_200PCS.xlsx'
wb = openpyxl.load_workbook(source_path)

# 获取源数据工作表（默认第一个）
ws_source = wb.active

# 获取FO提取工作表
if 'FO提取' in wb.sheetnames:
    ws_fo = wb['FO提取']
else:
    raise ValueError("工作簿中没有找到名为 'FO提取' 的工作表！")

# 获取源文件最大列数和行数
max_col = ws_source.max_column
max_row = ws_source.max_row

# 写入到FO提取工作表的起始位置
write_row = 1
write_col = 1
count_in_col = 0

# 从第2列开始处理
for col in range(2, max_col + 1):
    max_value = None
    candidate_rows = []

    # 找到这一列的最大值，以及对应的行号列表
    for row in range(1, max_row + 1):
        cell_value = ws_source.cell(row=row, column=col).value
        if isinstance(cell_value, (int, float)):
            if (max_value is None) or (cell_value > max_value):
                max_value = cell_value
                candidate_rows = [row]  # 新的最大值，重置列表
            elif cell_value == max_value:
                candidate_rows.append(row)  # 相同最大值，加入列表

    if candidate_rows:
        # 在candidate_rows中，找第一列最小的对应值
        min_first_col_value = None
        selected_row = None

        for row in candidate_rows:
            first_col_value = ws_source.cell(row=row, column=1).value
            if isinstance(first_col_value, (int, float)):
                if (min_first_col_value is None) or (first_col_value < min_first_col_value):
                    min_first_col_value = first_col_value
                    selected_row = row

        if selected_row is not None:
            corresponding_value = ws_source.cell(row=selected_row, column=1).value

            # 写入到FO提取表
            ws_fo.cell(row=write_row, column=write_col, value=corresponding_value)

            write_row += 1
            count_in_col += 1

            # 写满72个后换列
            if count_in_col >= 72:
                write_col += 1
                write_row = 1
                count_in_col = 0

# 保存
wb.save(source_path)

print("处理完成！已经写入FO提取工作表！")
