import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三个工作表，无表头
sheets = pd.read_excel(file_path, sheet_name=['THD', 'Fund', 'IMP'], header=None)

# 提取 DataFrame
df_thd, df_fund, df_imp = sheets['THD'], sheets['Fund'], sheets['IMP']

# 最大列索引（防止越界）
max_col_index = df_fund.shape[1] - 1

# 第23行（索引22）小于60的列，从第2列（索引1）开始
cols_23 = (df_fund.iloc[22, 1:] < 60)
drop_cols_23 = (cols_23[cols_23].index + 1).tolist()

# 第18行（索引17）小于61的列，从第2列（索引1）开始
cols_18 = (df_fund.iloc[17, 1:] < 61)
drop_cols_18 = (cols_18[cols_18].index + 1).tolist()

# 合并所有需要删除的列索引，去重并过滤越界
cols_to_drop = sorted(set(drop_cols_23 + drop_cols_18))
cols_to_drop = [col for col in cols_to_drop if col <= max_col_index]

# 删除对应列
df_fund.drop(columns=cols_to_drop, inplace=True)
df_thd.drop(columns=cols_to_drop, inplace=True)
df_imp.drop(columns=cols_to_drop, inplace=True)

# 保存修改
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    df_thd.to_excel(writer, sheet_name='THD', header=False, index=False)
    df_fund.to_excel(writer, sheet_name='Fund', header=False, index=False)
    df_imp.to_excel(writer, sheet_name='IMP', header=False, index=False)

# 输出结果
print(f"Fund 表第23行 <60 或 第18行 <61 的列已删除，共计：{len(cols_to_drop)} 列。")
