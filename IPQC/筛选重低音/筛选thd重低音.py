import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取3个工作表（无表头）
sheets = pd.read_excel(file_path, sheet_name=['THD', 'Fund', 'IMP'], header=None)

# 提取为 DataFrame
df_thd = sheets['THD']
df_fund = sheets['Fund']
df_imp = sheets['IMP']

# 初始化待删除列列表（从第2列开始，即索引1）
cols_to_drop = []

# 遍历列索引
for col_idx in range(1, df_thd.shape[1]):
    col_values = df_thd.iloc[21:100, col_idx]  # 第22~100行（索引21~99）
    if (col_values > 35).any():
        cols_to_drop.append(col_idx)

# 执行删除
df_thd.drop(columns=cols_to_drop, inplace=True)
df_fund.drop(columns=cols_to_drop, inplace=True)
df_imp.drop(columns=cols_to_drop, inplace=True)

# 保存回原文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    df_thd.to_excel(writer, sheet_name='THD', header=False, index=False)
    df_fund.to_excel(writer, sheet_name='Fund', header=False, index=False)
    df_imp.to_excel(writer, sheet_name='IMP', header=False, index=False)

# 输出删除列数量
print(f"THD 工作表中删除的列数：{len(cols_to_drop)}")
