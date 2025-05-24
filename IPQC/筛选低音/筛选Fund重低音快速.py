import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三个工作表，无表头
sheets = pd.read_excel(file_path, sheet_name=['THD', 'Fund', 'IMP'], header=None)

# 提取 DataFrame
df_thd, df_fund, df_imp = sheets['THD'], sheets['Fund'], sheets['IMP']

# 最大列索引
max_col_index = df_fund.shape[1] - 1

# 定义删除列条件函数
def get_drop_cols(row_idx, condition_fn, df):
    row_values = df.iloc[row_idx, 1:]  # 从第2列开始（索引1）
    mask = condition_fn(row_values)
    return (mask[mask].index + 1).tolist()  # 补偿索引偏移

# 获取所有需删除的列索引
drop_cols = set()

drop_cols.update(get_drop_cols(46, lambda x: x < 70, df_fund))   # 第23行 <60
# drop_cols.update(get_drop_cols(23, lambda x: x < 65, df_fund))   # 第24行 <65
# drop_cols.update(get_drop_cols(23, lambda x: x > 68, df_fund))   # 第24行 >68
# drop_cols.update(get_drop_cols(17, lambda x: x < 61, df_fund))   # 第18行 <61
# drop_cols.update(get_drop_cols(16, lambda x: x < 62, df_fund))   # 第17行 <62
# drop_cols.update(get_drop_cols(21, lambda x: x < 60, df_fund))   # 第22行 <60
# 过滤越界列并排序
cols_to_drop = sorted([col for col in drop_cols if col <= max_col_index])

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
print(f"已删除 Fund 中不满足条件的列，共 {len(cols_to_drop)} 列。")
