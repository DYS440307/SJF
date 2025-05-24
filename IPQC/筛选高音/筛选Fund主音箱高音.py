from openpyxl import load_workbook

# 加载工作簿
file_path = r"E:\System\pic\1.xlsx"
wb = load_workbook(file_path)

# 读取工作表
fund_ws = wb["Fund"]
thd_ws = wb["THD"]
imp_ws = wb["IMP"]

# 找出需要删除的列索引（从第2列开始）
cols_to_delete = []
for col in range(2, fund_ws.max_column + 1):
    cell_value = fund_ws.cell(row=48, column=col).value
    if isinstance(cell_value, (int, float)) and cell_value < 72:
        cols_to_delete.append(col)

# 删除列时要从后往前删，防止索引错乱
for col in reversed(cols_to_delete):
    fund_ws.delete_cols(col)
    thd_ws.delete_cols(col)
    imp_ws.delete_cols(col)

# 输出删除的列数
print(f"已从Fund中删除了 {len(cols_to_delete)} 列")

# 保存修改
wb.save(file_path)
