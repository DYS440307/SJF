import pandas as pd

# 文件路径
file_path = r"E:/System/pic/1.xlsx"

# 读取三个工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 初始化要删除的列索引列表
cols_to_delete = []

# 从第2列开始检查（索引为1开始）
for col in range(1, thd_df.shape[1]):
    col_values = thd_df.iloc[54:82, col]  # 对应 Excel 中第55~82行
    if (col_values > 10).any():
        cols_to_delete.append(col)

# 删除对应列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# 保存到原 Excel 文件中
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# 输出删除信息
print(f"THD 工作表中删除的列数：{len(cols_to_delete)}")
