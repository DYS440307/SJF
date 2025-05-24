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
    if pd.notna(col_values.iloc[0]) and col_values.iloc[1:93].isna().any():
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

# ========== 第二步：查找 Fund 中完全相同的重复列 ==========
duplicate_cols = []
seen = {}

for col in range(fund_df.shape[1]):
    col_key = tuple(fund_df.iloc[:, col].fillna('').astype(str))  # 转为字符串再转元组，用于精确比较
    if col_key in seen:
        duplicate_cols.append(col)  # 标记当前为重复列，删除
    else:
        seen[col_key] = col  # 第一次出现，记录下来

# ========== 合并所有需要删除的列 ==========
all_cols_to_delete = sorted(set(cols_to_delete + duplicate_cols))
# ⬇️ 删除 IMP / Fund / THD 中对应列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# ⬇️ 保存回原 Excel 文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# ===== 输出信息 =====
print(f"符合条件删除列数：{len(cols_to_delete)}")
print(f"重复列删除列数：{len(duplicate_cols)}")
print(f"总共删除列数：{len(all_cols_to_delete)}，列索引为：{all_cols_to_delete}")
