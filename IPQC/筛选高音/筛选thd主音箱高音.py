import pandas as pd

# 文件路径
file_path = r"E:/System/pic/1.xlsx"

# 读取三个工作表（不设置表头）
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df = sheets['IMP']
fund_df = sheets['Fund']
thd_df = sheets['THD']


# 设置删除条件的函数
def should_delete_column(thd_column: pd.Series, full_column: pd.Series) -> bool:
    """
    判断该列是否应删除，条件如下：
    1. THD 表中第55~82行中任意值大于10
    2. 首行有数值，且第2~82行有空值
    """

    # 条件1：首行非空，且第2~82行（索引1~81）存在空值
    condition2 = pd.notna(full_column.iloc[0]) and full_column.iloc[1:82].isna().any()
    # 条件2：THD工作表第55~82行（索引54~81）中是否有值大于10
    condition1 = (thd_column[54:82] > 10).any()

    return condition1 or condition2


# 记录要删除的列索引
cols_to_delete = []

# 遍历列（跳过第0列）
for col in range(1, thd_df.shape[1]):
    if should_delete_column(thd_df.iloc[:, col], thd_df.iloc[:, col]):
        cols_to_delete.append(col)

# 删除这些列
imp_df.drop(columns=cols_to_delete, inplace=True)
fund_df.drop(columns=cols_to_delete, inplace=True)
thd_df.drop(columns=cols_to_delete, inplace=True)

# 保存到 Excel 文件中
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

# 输出信息
print(f"删除列索引：{cols_to_delete}")
print(f"总共删除了 {len(cols_to_delete)} 列。")
