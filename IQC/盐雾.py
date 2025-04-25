import warnings
from openpyxl import load_workbook

# 抑制 openpyxl 对图表读取时的警告
warnings.filterwarnings("ignore", ".*Unable to read chart.*", UserWarning)

# 文件路径
file1 = r"E:\System\desktop\2025年.xlsx"
file2 = r"E:\System\desktop\盐雾实验记录.xlsx"

# 要操作的工作表名称（例如"1月"）
sheet_name = "1月"

# 关键词列表
keywords = ['T铁', 'U铁', '盆架', '钕铁硼', '华司']

# 加载工作簿与指定工作表
wb1 = load_workbook(filename=file1, data_only=True)
sheet1 = wb1[sheet_name]  # 从名称为"1月"的工作表读取

wb2 = load_workbook(filename=file2)
sheet2 = wb2[sheet_name]  # 写入到名称相同的工作表

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
print(f"数据迁移完成：共迁移 {current_row - start_row} 行符合条件的记录，写入到盐雾实验记录.xlsx 的 '{sheet_name}' 表，从第{start_row}行开始。")
