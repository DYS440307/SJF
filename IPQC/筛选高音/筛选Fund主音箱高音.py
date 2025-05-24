import pandas as pd
from openpyxl import load_workbook

# 文件路径
file_path = r"E:\System\pic\1.xlsx"

# 使用pandas读取Fund工作表
df_fund = pd.read_excel(file_path, sheet_name="Fund", engine="openpyxl", header=None)

# 找出第48行（索引47）中值小于72的列索引（从第2列开始，索引为1）
cols_to_delete = [col for col in range(1, df_fund.shape[1])  # 从第2列开始
                  if pd.api.types.is_number(df_fund.iloc[47, col]) and df_fund.iloc[47, col] < 72]

# 打印将要删除的列数
print(f"已从Fund中删除了 {len(cols_to_delete)} 列")

# 用 openpyxl 删除3张表中对应列
wb = load_workbook(file_path)
for sheet_name in ["Fund", "THD", "IMP"]:
    ws = wb[sheet_name]
    for col in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col + 1)  # openpyxl是从1开始的列索引，所以加1

# 保存文件
wb.save(file_path)
