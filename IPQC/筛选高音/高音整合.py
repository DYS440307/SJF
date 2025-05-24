import pandas as pd

# 文件路径
file_path = 'E:/System/pic/高音.xlsx'

# 读取三个工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']

# ===== 定义各表的筛选函数 =====

def should_delete_imp_column(col_values: pd.Series) -> bool:
    if pd.notna(col_values.iloc[0]) and col_values.iloc[0:93].isna().any():
        return True
    if (pd.to_numeric(col_values.iloc[0:94], errors='coerce') > 9).any():
        return True
    if (pd.to_numeric(col_values.iloc[63:66], errors='coerce') > 6.1).any():
        return True
    if (pd.to_numeric(col_values.iloc[45:47], errors='coerce') < 6.1).any():
        return True
    if (pd.to_numeric(col_values.iloc[36:39], errors='coerce') < 5.78).any():
        return True
    if (pd.to_numeric(col_values.iloc[0:10], errors='coerce') > 5.9).any():
        return True
    return False

def should_delete_fund_column(col_values: pd.Series) -> bool:
    numeric_values = pd.to_numeric(col_values, errors='coerce')
    if pd.notna(col_values.iloc[0]) and col_values.iloc[1:93].isna().any():
        return True
    if (numeric_values.iloc[50:54] < 85).any():
        return True
    return False

def should_delete_thd_column(col_values: pd.Series) -> bool:
    numeric_values = pd.to_numeric(col_values, errors='coerce')
    if pd.notna(col_values.iloc[0]) and col_values.iloc[0:81].isna().any():
        return True
    if (numeric_values.iloc[24:81] > 15).any():
        return True
    return False

# ===== 判断每张表需删除的列（从第2列开始） =====

imp_delete = [col for col in range(1, imp_df.shape[1]) if should_delete_imp_column(imp_df.iloc[:, col])]
fund_delete = [col for col in range(1, fund_df.shape[1]) if should_delete_fund_column(fund_df.iloc[:, col])]
thd_delete = [col for col in range(1, thd_df.shape[1]) if should_delete_thd_column(thd_df.iloc[:, col])]

# ===== 检查重复列（每张表分别判断） =====

def find_duplicate_columns(df: pd.DataFrame) -> list:
    seen = {}
    dup = []
    for col in range(df.shape[1]):
        col_tuple = tuple(df.iloc[:, col].fillna('').astype(str))
        if col_tuple in seen:
            dup.append(col)
        else:
            seen[col_tuple] = col
    return dup

imp_dup = find_duplicate_columns(imp_df)
fund_dup = find_duplicate_columns(fund_df)
thd_dup = find_duplicate_columns(thd_df)

# ===== 合并所有需删除的列索引 =====
all_delete_cols = sorted(set(imp_delete + fund_delete + thd_delete + imp_dup + fund_dup + thd_dup))

# ===== 删除 IMP/Fund/THD 中对应列 =====
imp_df.drop(columns=all_delete_cols, inplace=True)
fund_df.drop(columns=all_delete_cols, inplace=True)
thd_df.drop(columns=all_delete_cols, inplace=True)

# ===== 写回 Excel 文件 =====
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# ===== 输出信息 =====
print("IMP 删除列：", len(imp_delete), " 重复列：", len(imp_dup))
print("Fund 删除列：", len(fund_delete), " 重复列：", len(fund_dup))
print("THD 删除列：", len(thd_delete), " 重复列：", len(thd_dup))
print(f"总共删除列数：{len(all_delete_cols)}，列索引为：{all_delete_cols}")
