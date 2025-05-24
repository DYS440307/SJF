import pandas as pd

# 文件路径
file_path = r"E:/System/pic/1.xlsx"

# 读取 Excel 文件的3个工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 找出第48行（索引47）中值小于72的列索引（从第2列开始，索引为1）
cols_to_delete = []

for col in range(1, fund_df.shape[1]):
    val = fund_df.iloc[47, col]
    if pd.notna(val) and isinstance(val, (int, float)) and val < 72:
        cols_to_delete.append(col)

# 删除列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# 写入原文件（使用 'w' 模式覆盖保存）
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

print(f"已从Fund中删除了 {len(cols_to_delete)} 列")
