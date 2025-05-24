import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三个工作表，无表头
sheets = pd.read_excel(file_path, sheet_name=['THD', 'Fund', 'IMP'], header=None)

# 提取 DataFrame
df_thd, df_fund, df_imp = sheets['THD'], sheets['Fund'], sheets['IMP']

# 提取要筛选的行范围（第22~100行，索引21~99），以及要判断的列范围（从第1列开始）
data_to_check = df_thd.iloc[21:100, 1:]

# ----------- 筛选逻辑开始（可修改） -----------

# 筛选条件：某列中是否存在大于35的值
condition = data_to_check > 35

# 符合条件的列（布尔Series）：任意一行满足条件即为True
columns_meeting_condition = condition.any()

# 获取这些列的原始索引（因为data_to_check从第1列开始，要加1）
cols_to_drop = (columns_meeting_condition[columns_meeting_condition].index + 1).tolist()

# ----------- 筛选逻辑结束 -----------

# 删除对应列
df_thd.drop(columns=cols_to_drop, inplace=True)
df_fund.drop(columns=cols_to_drop, inplace=True)
df_imp.drop(columns=cols_to_drop, inplace=True)

# 保存到原文件
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    df_thd.to_excel(writer, sheet_name='THD', header=False, index=False)
    df_fund.to_excel(writer, sheet_name='Fund', header=False, index=False)
    df_imp.to_excel(writer, sheet_name='IMP', header=False, index=False)

# 输出删除列数量
print(f"THD 工作表中删除的列数：{len(cols_to_drop)}")
