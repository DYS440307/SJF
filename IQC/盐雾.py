from openpyxl import load_workbook

# 文件路径
file1 = r"E:\System\desktop\2025年.xlsx"
file2 = r"E:\System\desktop\盐雾实验记录.xlsx"

# 关键词列表
keywords = ['T铁', 'U铁', '盆架', '钕铁硼', '华司', '端子板']

# 加载工作簿与工作表
wb1 = load_workbook(filename=file1)
sheet1 = wb1.active  # 源表，假设数据在第一个工作表

wb2 = load_workbook(filename=file2)
sheet2 = wb2.active  # 目标表，假设写入第一个工作表

# 从目标表的第4行开始写入
start_row = 4
current_row = start_row

# 遍历源表从第2行开始的所有数据行
for row in sheet1.iter_rows(min_row=2, values_only=False):
    cell_value = row[2].value  # 第三列数据
    # 判断是否包含任意关键词
    if cell_value and any(kw in str(cell_value) for kw in keywords):
        # 将前四列的值写入目标表对应行的 A-D 列
        for col_idx in range(4):
            target_cell = sheet2.cell(row=current_row, column=col_idx + 1)
            target_cell.value = row[col_idx].value
        current_row += 1  # 写完一行后，目标行号加1

# 保存目标工作簿
wb2.save(filename=file2)
print(f"数据迁移完成：共迁移 {current_row - start_row} 行符合条件的记录，写入到盐雾实验记录.xlsx 的第{start_row}行开始。")
