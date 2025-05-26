import pandas as pd

# 文件路径
file_path = 'E:\System\pic\全\低音\原档筛选.xlsx'

# 读取三个工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# 设置删除条件的函数
def should_delete_column(thd_column: pd.Series) -> bool:
    """
    判断该列是否应删除，满足任一条件返回 True：
    1. THD 表中第55~82行（索引54~81）中任意值 > 10
    2. 首行有值，且第2~82行（索引1~81）中存在空值
    """
    thd_numeric = pd.to_numeric(thd_column, errors='coerce')

    # 条件1：首行有值，且第2~82行有空值
    if pd.notna(thd_column.iloc[0]) and thd_column.iloc[1:82].isna().any():
        return True

    # 条件2：第55~82行（索引54~81）任意值 > 10
    if (thd_numeric.iloc[28:109] > 20).any():
        return True


    return False

# 记录需删除列索引（从第2列开始）
cols_to_delete = []

for col in range(1, thd_df.shape[1]):
    if should_delete_column(thd_df.iloc[:, col]):
        cols_to_delete.append(col)

# ===== 查找 THD 中完全相同的列（保留第一列） =====
duplicate_cols = []
seen = {}

for col in range(thd_df.shape[1]):
    col_tuple = tuple(thd_df.iloc[:, col].fillna('').astype(str))
    if col_tuple in seen:
        duplicate_cols.append(col)
    else:
        seen[col_tuple] = col

# ===== 合并所有要删除的列索引 =====
all_cols_to_delete = sorted(set(cols_to_delete + duplicate_cols))

# 删除 IMP / Fund / THD 中对应列
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
