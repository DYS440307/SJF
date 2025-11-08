import pandas as pd

# 文件路径
file_path = r"E:\System\download\收料通知单_2025110720565501_236281 - 副本.xlsx"

# 读取数据
df = pd.read_excel(file_path)

# 填充物料编码空白（有时某些行会缺）
df['物料编码'] = df['物料编码'].ffill()

# 去重（物料编码+供应商）
df = df.drop_duplicates(subset=['物料编码', '供应商'], keep='first')

# 按物料编码分组合并供应商，用 ; 分隔
merged_df = df.groupby('物料编码', as_index=False).agg({'供应商': lambda x: ';'.join(map(str, x))})

# 覆盖保存
merged_df.to_excel(file_path, index=False)

print("✅ 已完成：物料编码下合并供应商，并覆盖原文件。")
