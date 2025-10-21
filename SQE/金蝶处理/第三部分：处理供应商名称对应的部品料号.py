import pandas as pd

# 文件路径
file_path = r"E:\System\desktop\PY\图纸归档系统\供应商原档 - 副本.xlsx"

# 读取数据
df = pd.read_excel(file_path)

# 填充供应商空白
df['供应商'] = df['供应商'].ffill()

# 去重（供应商+物料编码）
df = df.drop_duplicates(subset=['供应商', '物料编码'], keep='first')

# 按供应商分组合并物料编码，用 ; 分隔
merged_df = df.groupby('供应商', as_index=False).agg({'物料编码': lambda x: ';'.join(map(str, x))})

# 覆盖保存
merged_df.to_excel(file_path, index=False)

print("✅ 已完成：供应商下合并物料编码，并覆盖原文件。")
