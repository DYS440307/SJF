import pandas as pd

# Excel 文件路径
file_path = 'E:/System/pic/1.xlsx'

# 读取三个工作表，无表头
sheets = pd.read_excel(file_path, sheet_name=['THD', 'Fund', 'IMP'], header=None)

# 提取 DataFrame
df_thd, df_fund, df_imp = sheets['THD'], sheets['Fund'], sheets['IMP']

# 设置检查范围：第55~82行（索引54~82），第2列开始（索引1开始）
row_start, row_end = 57, 75  # iloc 切片末尾不包含
col_start = 1

# 获取对应数据范围
data_to_check = df_thd.iloc[row_start:row_end, col_start:]

# 条件：是否存在大于20
condition = data_to_check > 20

# 找出满足条件的列索引（注意偏移）
cols_to_drop = (condition.any()).loc[lambda x: x].index + col_start

# 转为列表并去重排序
cols_to_drop = sorted(set(cols_to_drop.tolist()))

# 安全过滤，防止越界
max_col = df_thd.shape[1] - 1
cols_to_drop = [col for col in cols_to_drop if col <= max_col]

# 执行删除
df_thd.drop(columns=cols_to_drop, inplace=True)
df_fund.drop(columns=cols_to_drop, inplace=True)
df_imp.drop(columns=cols_to_drop, inplace=True)

# 保存修改
with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
    df_thd.to_excel(writer, sheet_name='THD', header=False, index=False)
    df_fund.to_excel(writer, sheet_name='Fund', header=False, index=False)
    df_imp.to_excel(writer, sheet_name='IMP', header=False, index=False)

# 输出信息
print(f"从 THD 第{row_start+1}~{row_end}行中删除大于20的列，共删除 {len(cols_to_drop)} 列。")
