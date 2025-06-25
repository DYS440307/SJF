import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import time

# 配置参数 - 可直接修改以下值
TARGET_VALUE = 600
FILE_PATH = r"E:\System\pic\A报告\IMP数据.xlsx"


def find_nearest_value(df, target):
    left, right = 0, len(df) - 1
    nearest_idx = 0
    min_diff = float('inf')

    while left <= right:
        mid = (left + right) // 2
        current_diff = abs(df.iloc[mid, 0] - target)

        if current_diff < min_diff:
            min_diff = current_diff
            nearest_idx = mid

        if df.iloc[mid, 0] < target:
            left = mid + 1
        elif df.iloc[mid, 0] > target:
            right = mid - 1
        else:
            break

    return df.iloc[nearest_idx, 1]


try:
    start_time = time.time()

    # 读取Excel文件
    excel_file = pd.ExcelFile(FILE_PATH)

    # 获取IMP原档中的数据
    df = excel_file.parse("IMP原档")

    # 检查IMP原档中有数值的列数是否为偶数
    non_empty_columns = df.dropna(axis=1, how='all').shape[1]

    if non_empty_columns % 2 != 0:
        raise ValueError(f"IMP原档中有数值的列数为{non_empty_columns}，必须为偶数")

    # 使用openpyxl将值写入ACR表
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb["ACR"]

    # 处理所有奇数列(1,3,5...)及其后续偶数列(2,4,6...)
    max_col = df.shape[1]
    for col_pair in range(0, max_col, 2):
        # 检查是否有足够的列
        if col_pair + 1 >= max_col:
            break

        # 获取当前奇数列和偶数列
        odd_col = col_pair
        even_col = col_pair + 1

        # 提取当前奇数列和偶数列的数据
        current_df = df.iloc[:, [odd_col, even_col]].dropna()

        # 跳过空列对
        if current_df.empty:
            continue

        # 找到奇数列最接近目标值的值对应的偶数列的值
        nearest_value = find_nearest_value(current_df, TARGET_VALUE)

        # 计算ACR表中的行号(从1开始，每对占一行)
        row_in_acr = (col_pair // 2) + 1

        # 将值写入ACR表的对应行的第一列
        ws.cell(row=row_in_acr, column=1).value = nearest_value

    # 保存修改
    wb.save(FILE_PATH)

    end_time = time.time()
    execution_time = end_time - start_time

    print(f"已成功处理所有列对数据并写入'ACR'表")
    print(f"程序运行时间: {execution_time:.4f} 秒")

except Exception as e:
    print(f"发生错误: {e}")