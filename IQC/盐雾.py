import pandas as pd

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

# 写入新的Excel文件
output_path = r'F:\system\Desktop\PY\IQC\盐雾实验记录.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_filtered.to_excel(writer, index=False, header=False, startrow=3, sheet_name='Sheet1')
