import pandas as pd

# 文件路径
file_path = 'E:\System\pic\全\低音\原档筛选.xlsx'

# 读取三张工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 用于判断某列是否应删除的函数
def should_delete_column(col_values: pd.Series) -> bool:

    if pd.notna(col_values.iloc[0]) and col_values.iloc[1:94].isna().any():
        return True
    if (pd.to_numeric(col_values.iloc[0:13], errors='coerce') > 9).any():
        return True
    if (pd.to_numeric(col_values.iloc[32:68], errors='coerce') >10).any():
        return True
    if (pd.to_numeric(col_values.iloc[84:97], errors='coerce') > 20).any():
        return True
    if (pd.to_numeric(col_values.iloc[0:109], errors='coerce') > 40).any():
        return True
    if (pd.to_numeric(col_values.iloc[23:24], errors='coerce') < 15).any():
        return True
    if (pd.to_numeric(col_values.iloc[116:118], errors='coerce') > 50).any():
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
