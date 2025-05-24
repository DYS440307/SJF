import pandas as pd

# 文件路径
file_path = r"E:/System/pic/1.xlsx"

# 读取 Excel 文件的3个工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# ⬇️ 用于判断 Fund 中某列是否应删除
def should_delete_column(col_values: pd.Series) -> bool:
    """
    判断该列是否应删除，满足任一条件即返回 True
    """
    # 安全转换为数值类型
    numeric_values = pd.to_numeric(col_values, errors='coerce')

    # 条件1：首行有值，且第2~94行有空值（索引0有值，索引1-93中有空）
    if pd.notna(col_values.iloc[0]) and col_values.iloc[1:94].isna().any():
        return True

    # 条件2：第1~94行中任一值 > 9
    if (numeric_values.iloc[50:54] < 85).any():
        return True


    return False

# ⬇️ 执行筛选并记录要删除的列（从第2列开始）
cols_to_delete = []
for col in range(1, fund_df.shape[1]):
    col_values = fund_df.iloc[:, col]
    if should_delete_column(col_values):
        cols_to_delete.append(col)

# ⬇️ 删除 IMP / Fund / THD 中对应列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# ⬇️ 保存回原 Excel 文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

print(f"已从 Fund 中删除了 {len(cols_to_delete)} 列")
