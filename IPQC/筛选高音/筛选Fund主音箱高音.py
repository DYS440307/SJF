import pandas as pd

# 文件路径
file_path = 'E:\System\pic\全\高音\原档筛选.xlsx'

# 读取 Excel 文件的 3 个工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 判断 Fund 中某列是否应删除
def should_delete_column(col_values: pd.Series) -> bool:
    """
    判断该列是否应删除，满足任一条件即返回 True：
    1. 第1行有值，且第2~94行中有空值；
    2. 第51~54行中存在小于85的值；
    """
    numeric_values = pd.to_numeric(col_values, errors='coerce')

    # 条件1：第1行有值，且第2~94行有空值
    if pd.notna(col_values.iloc[0]) and col_values.iloc[1:94].isna().any():
        return True

    # 条件2：第51~54行中任意值小于 85
    if (numeric_values.iloc[50:54] < 85).any():
        return True

    return False

# 查找不合格列（从第2列开始）
cols_to_delete = []
for col in range(1, fund_df.shape[1]):
    if should_delete_column(fund_df.iloc[:, col]):
        cols_to_delete.append(col)

# 查找完全重复的列
duplicate_cols = []
seen = {}
for col in range(fund_df.shape[1]):
    col_key = tuple(fund_df.iloc[:, col].fillna('').astype(str))
    if col_key in seen:
        duplicate_cols.append(col)
    else:
        seen[col_key] = col

# 合并所有需删除列并排序
all_cols_to_delete = sorted(set(cols_to_delete + duplicate_cols))

# 删除 IMP / Fund / THD 中对应列
imp_df.drop(columns=all_cols_to_delete, inplace=True)
fund_df.drop(columns=all_cols_to_delete, inplace=True)
thd_df.drop(columns=all_cols_to_delete, inplace=True)

# 保存回原 Excel 文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# 输出信息
print(f"符合条件删除列数：{len(cols_to_delete)}")
print(f"重复列删除列数：{len(duplicate_cols)}")
print(f"总共删除列数：{len(all_cols_to_delete)}，列索引为：{all_cols_to_delete}")
