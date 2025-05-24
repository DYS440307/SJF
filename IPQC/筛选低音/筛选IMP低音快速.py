import pandas as pd

# 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取 Excel 文件中的多个工作表
sheets = pd.read_excel(file_path, sheet_name=['IMP', 'Fund', 'THD'], header=None)
imp_df, fund_df, thd_df = sheets['IMP'], sheets['Fund'], sheets['THD']

# 定义列删除条件的 lambda 函数
should_delete = lambda col: (
    (col[1:93] > 20).any() or
    (pd.notna(col.iloc[0]) and col.isna().any())  # 新增条件：首行有值且该列有空值
     or  (col[1:28] > 10).any()
    or (col[46:55] > 15.5).any()
    # 如果你想启用以下条件，请取消注释
    # or (col[14:27] < 5).any()
    # or (col[2:110] > 40).any()
    # or (col[99:112] > 30).any()
)

# 获取要删除的列索引（从第2列开始）
cols_to_delete = [col for col in range(1, imp_df.shape[1]) if should_delete(imp_df.iloc[:, col])]

# 删除对应列
for df in [imp_df, fund_df, thd_df]:
    df.drop(columns=cols_to_delete, inplace=True)

# 保存修改后的工作表
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    imp_df.to_excel(writer, sheet_name='IMP', header=False, index=False)
    fund_df.to_excel(writer, sheet_name='Fund', header=False, index=False)
    thd_df.to_excel(writer, sheet_name='THD', header=False, index=False)

print(f"总共删除了 {len(cols_to_delete)} 列。")
