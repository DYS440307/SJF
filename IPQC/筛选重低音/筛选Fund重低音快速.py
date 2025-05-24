import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三个工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)

# 提取各工作表
df_imp = sheets['IMP']
df_fund = sheets['Fund']
df_thd = sheets['THD']

# 要删除的列索引（从第2列开始，即索引为1）
cols_to_drop = []

# 遍历每一列（从第2列开始）
for col_idx in range(1, df_fund.shape[1]):
    delete_flag = False  # 删除标志，默认不删除

    # 条件1：第17行（索引16）值小于60
    if df_fund.iloc[16, col_idx] < 59:
        delete_flag = True

        # 条件1：第17行（索引16）值小于60
    if df_fund.iloc[24, col_idx] > 69:
            delete_flag = True


    # 条件2：第23~27行（索引22~26）中有任意值小于60
    if (df_fund.iloc[22:27, col_idx] < 60).any():
        delete_flag = True

    # 如果满足任一条件，则标记删除
    if delete_flag:
        cols_to_drop.append(col_idx)

# 删除列
df_imp.drop(columns=cols_to_drop, inplace=True)
df_fund.drop(columns=cols_to_drop, inplace=True)
df_thd.drop(columns=cols_to_drop, inplace=True)

# 写入原 Excel 文件（覆盖保存）
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    df_imp.to_excel(writer, sheet_name='IMP', header=False, index=False)
    df_fund.to_excel(writer, sheet_name='Fund', header=False, index=False)
    df_thd.to_excel(writer, sheet_name='THD', header=False, index=False)

print(f"总共删除了 {len(cols_to_drop)} 列。")
