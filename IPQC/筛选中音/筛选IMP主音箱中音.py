import pandas as pd

# 文件路径
file_path = 'E:\System\pic\全\中音\筛选.xlsx'

# 读取三张工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 用于判断某列是否应删除的函数
def should_delete_column(col_values: pd.Series) -> bool:
    """
    判断该列是否应删除，满足任一条件即返回 True
    """
    # 条件1（新增）：首行有值，且第2~94行有空值
    if pd.notna(col_values.iloc[0]) and col_values.iloc[0:93].isna().any():
        return True
    # 条件2：
    if (pd.to_numeric(col_values.iloc[0:93], errors='coerce') > 17).any():
        return True
    # 条件3：
    if (pd.to_numeric(col_values.iloc[81:93], errors='coerce') > 10).any():
        return True
    return False

# 存储需要删除的列索引（从第2列开始）
cols_to_delete = []
for col in range(1, imp_df.shape[1]):
    if should_delete_column(imp_df.iloc[:, col]):
        cols_to_delete.append(col)

# 查找重复列（整列值一致）
duplicate_cols = []
seen = {}
for col in range(imp_df.shape[1]):
    col_tuple = tuple(imp_df.iloc[:, col].fillna('').astype(str))
    if col_tuple in seen:
        duplicate_cols.append(col)
    else:
        seen[col_tuple] = col

# 合并并去重排序所有待删除列
all_cols_to_delete = sorted(set(cols_to_delete + duplicate_cols))

# 删除对应列
imp_df.drop(columns=all_cols_to_delete, inplace=True)
fund_df.drop(columns=all_cols_to_delete, inplace=True)
thd_df.drop(columns=all_cols_to_delete, inplace=True)

# 写入原 Excel 文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# 输出信息
print(f"符合条件删除列数：{len(cols_to_delete)}")
print(f"重复列删除列数：{len(duplicate_cols)}")
print(f"总共删除列数：{len(all_cols_to_delete)}，列索引为：{all_cols_to_delete}")
