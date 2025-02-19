import pandas as pd
from openpyxl import load_workbook

# 读取Excel文件
file_path = r'F:\system\Desktop\PY\IQC\2025年.xlsx'
df = pd.read_excel(file_path, sheet_name='2月')

# 筛选部品类型
valid_types = ['T铁', 'U铁', '盆架', '钕铁硼', '华司']
df_filtered = df[df['部品类型'].isin(valid_types)].copy()  # 使用.copy()确保是副本

# 去重：同一月份内，相同供应商和相同料号的部品只保留一次
df_filtered['日期'] = pd.to_datetime(df_filtered['日期'])
df_filtered['月份'] = df_filtered['日期'].dt.month

# 去重操作
df_filtered = df_filtered.drop_duplicates(subset=['月份', '供应商', '料号'])

# 重新排列列顺序：将日期列排在第一列，其他列后面
df_filtered = df_filtered[['日期', '供应商', '部品类型', '料号']]  # 日期排第一

# 读取现有的Excel文件（包括目标表格中的内容）
output_path = r'F:\system\Desktop\PY\IQC\盐雾实验记录.xlsx'

# 使用 openpyxl 加载现有工作簿
book = load_workbook(output_path)

# 获取第一个工作表
sheet = book.worksheets[0]  # 第一个工作表

# 获取现有工作表的前三行数据
existing_data = pd.read_excel(output_path, sheet_name=sheet.title, header=None)

# 获取当前工作表的总行数
existing_row_count = existing_data.shape[0]

# 从第四行开始写入新数据
for idx, row in enumerate(df_filtered.values, start=existing_row_count + 1):
    target_row = idx + 1  # 目标行数，从第四行开始

    # 插入日期、供应商、部品类型和料号到当前行
    sheet.append(row.tolist())

    # 检查日期列（第一列）是否有数据，如果有，填入第五到第八列的固定值
    if row[0]:  # 如果日期列有数据
        # 填写第五列到第八列
        sheet.cell(row=target_row, column=5, value="5PCS")  # 第五列写入 5PCS
        sheet.cell(row=target_row, column=6, value="无")    # 第六列写入 无
        sheet.cell(row=target_row, column=7, value="合格")  # 第七列写入 合格
        sheet.cell(row=target_row, column=8, value="邓洋枢")  # 第八列写入 邓洋枢

# 保存修改后的工作簿
book.save(output_path)
