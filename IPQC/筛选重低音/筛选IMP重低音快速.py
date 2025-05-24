import pandas as pd

# 读取 Excel 文件的3个工作表
file_path = 'E:/System/pic/1.xlsx'
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)

imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 要删除的列索引（从第2列开始，即索引为1）
cols_to_delete = []

# 遍历所有列（从第2列开始）
for col in range(1, imp_df.shape[1]):
    col_values = imp_df.iloc[:, col]

    # 条件1：第2~14行有值 > 7
    if (col_values.iloc[1:14] > 7).any():
        cols_to_delete.append(col)
        continue

    # 条件2a：第39~74行有值 > 10
    if (col_values.iloc[38:74] > 10).any():
        cols_to_delete.append(col)
        continue

    # 条件2b：第15~27行有值 < 5
    if (col_values.iloc[14:27] < 5).any():
        cols_to_delete.append(col)
        continue

        # 条件2b：第15~27行有值 < 5
    if (col_values.iloc[2:110] > 40).any():
            cols_to_delete.append(col)
            continue

        # 条件2b：第15~27行有值 < 5
    if (col_values.iloc[99:112] > 30).any():
            cols_to_delete.append(col)
            continue

# 删除列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# 写入原文件（使用 'w' 模式覆盖保存）
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)


print(f"总共删除了 {len(cols_to_delete)} 列。")
