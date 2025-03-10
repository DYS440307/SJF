import openpyxl

# 打开工作簿和工作表
filepath = r'F:\system\Downloads\123.xlsx'
wb = openpyxl.load_workbook(filepath)
ws = wb.active

col = 2  # 从偶数列开始（B列）
while col <= ws.max_column:
    delete_flag = False
    for row in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=col).value
        try:
            if cell_value is not None:
                value = float(cell_value)
                if value < 60 or value > 90:
                    delete_flag = True
                    break
        except ValueError:
            continue  # 非数值忽略

    if delete_flag:
        ws.delete_cols(col - 1, 2)  # 删除奇数列和偶数列
        col = col - 2  # 向前回退两列，以适配删除后的列变动
        if col < 2:
            col = 2
    else:
        col += 2  # 没有删除则继续处理下一个偶数列

# 保存工作簿
wb.save(filepath)
print("处理完成，已删除不符合条件的列组。")
