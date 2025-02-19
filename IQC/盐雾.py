import pandas as pd
from openpyxl import load_workbook

# 读取Excel文件
file_path = r'F:\system\Desktop\PY\IQC\2025年.xlsx'
df = pd.read_excel(file_path, sheet_name='1月')

# 筛选部品类型
valid_types = ['T铁', 'U铁', '盆架', '钕铁硼', '华司']
df_filtered = df[df['部品类型'].isin(valid_types)].copy()  # 使用.copy()确保是副本

# 去重：同一月份内，相同供应商和相同料号的部品只保留一次
df_filtered['日期'] = pd.to_datetime(df_filtered['日期'])
df_filtered['月份'] = df_filtered['日期'].dt.month

# 去重操作
df_filtered = df_filtered.drop_duplicates(subset=['月份', '供应商', '料号'])

# 只选择需要的列：第一列到第四列（日期、供应商、部品类型、料号）
df_filtered = df_filtered[['日期', '供应商', '部品类型', '料号']]

# 读取现有的Excel文件（包括目标表格中的内容）
output_path = r'F:\system\Desktop\PY\IQC\盐雾实验记录.xlsx'

# 使用 openpyxl 加载现有工作簿
book = load_workbook(output_path)

# 检查是否存在目标工作表，并加载该工作表
if 'Sheet1' in book.sheetnames:
    sheet = book['Sheet1']
else:
    # 如果目标工作表不存在，则创建一个新的工作表
    sheet = book.create_sheet('Sheet1')

# 获取现有工作表的前三行数据
existing_data = pd.read_excel(output_path, sheet_name='Sheet1', header=None)

# 获取当前工作表的总行数
existing_row_count = existing_data.shape[0]

# 如果目标工作表中已经有数据，从第四行开始写入新数据
for idx, row in enumerate(df_filtered.values, start=existing_row_count + 1):
    # 在现有数据下方追加新数据
    sheet.append(row.tolist())

# 保存修改后的工作簿
book.save(output_path)
