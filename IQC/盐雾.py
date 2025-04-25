from openpyxl import load_workbook

# 文件路径
file1 = r"E:\System\desktop\2025年.xlsx"
file2 = r"E:\System\desktop\盐雾实验记录.xlsx"

# 关键词列表
keywords = ['T铁', 'U铁', '盆架', '钕铁硼', '华司']

# 加载工作簿与工作表
wb1 = load_workbook(filename=file1)
sheet1 = wb1.active  # 假设数据在第一个工作表中

wb2 = load_workbook(filename=file2)
sheet2 = wb2.active  # 假设目标数据写入到第一个工作表中

# 搜索符合条件的行并转移数据
for row in sheet1.iter_rows(min_row=2, values_only=False):  # 跳过表头，从第二行开始
    cell_value = row[2].value  # 第三列
    if cell_value and any(kw in str(cell_value) for kw in keywords):
        # 找到匹配行，提取前四列
        values_to_copy = [row[i].value for i in range(4)]
        # 写入到目标工作表指定单元格
        sheet2['A4'], sheet2['B4'], sheet2['C4'], sheet2['D4'] = values_to_copy
        break  # 如果只需要第一行匹配，找到后退出循环

# 保存目标工作簿
wb2.save(filename=file2)
print("数据迁移完成：已将符合条件的行前四列写入到盐雾实验记录.xlsx 的 A4:D4。")
