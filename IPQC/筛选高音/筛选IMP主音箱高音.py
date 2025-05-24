import pandas as pd

# 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三张工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 存储需要删除的列索引（从第2列开始）
cols_to_delete = []

for col in range(1, imp_df.shape[1]):  # 从第2列开始，index = 1
    delete_flag = False
    col_values = imp_df.iloc[:, col]

    # 条件1：第2~42行（Excel第2~43行）有值 > 7
    if (col_values.iloc[1:43] > 7).any():
        delete_flag = True

    # 条件2：第62~74行有值 > 6.3
    elif (col_values.iloc[61:74] > 6.3).any():
        delete_flag = True

    # 条件3a：第22~37行有值 > 6
    elif (col_values.iloc[21:37] > 6).any():
        delete_flag = True

    # 条件3b：第22~37行有值 < 5.55
    elif (col_values.iloc[21:37] < 5.55).any():
        delete_flag = True

    # 条件4：第2~93行（即索引1~93）中有值 > 10
    elif (col_values.iloc[1:94] > 10).any():
        delete_flag = True

    if delete_flag:
        cols_to_delete.append(col)

# 删除对应列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# 写入原文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# 输出删除列数
print(f"总共删除了 {len(cols_to_delete)} 列。")
