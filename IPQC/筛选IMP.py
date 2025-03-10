import openpyxl

# 打开工作簿和工作表
filepath = r'F:\system\Downloads\1.xlsx'
wb = openpyxl.load_workbook(filepath)
ws = wb.active

col = 2  # 从第2列开始（B列）
while col <= ws.max_column:
    delete_flag = False
    # 检查该偶数列的前40行
    for row in range(1, min(41, ws.max_row + 1)):
        cell_value = ws.cell(row=row, column=col).value
        try:
            if cell_value is not None and float(cell_value) > 20:
                delete_flag = True
                break
        except ValueError:
            continue

    if delete_flag:
        # 删除对应的奇数列（col-1）和偶数列（col）
        ws.delete_cols(col - 1, 2)  # 一次删除两列
        col = col - 2  # 回退两列重新判断当前位置
        if col < 2:
            col = 2
    else:
        col += 2  # 没删就跳到下一个偶数列

# 保存文件
wb.save(filepath)
print("处理完成，符合条件的列已删除。")
