import pandas as pd
import os
import time


def find_closest_value(sorted_series, target=600):
    """使用二分查找在排序后的Series中找到最接近目标值的索引"""
    left, right = 0, len(sorted_series) - 1
    closest_idx = left

    while left <= right:
        mid = (left + right) // 2

        # 更新最接近的值
        if abs(sorted_series.iloc[mid] - target) < abs(sorted_series.iloc[closest_idx] - target):
            closest_idx = mid

        # 继续搜索
        if sorted_series.iloc[mid] < target:
            left = mid + 1
        elif sorted_series.iloc[mid] > target:
            right = mid - 1
        else:
            # 找到完全相等的值，直接返回
            return mid

    return closest_idx


def process_excel_file(file_path):
    start_time = time.time()  # 开始计时
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        # 使用ExcelFile类一次性读取文件，提高性能
        excel_file = pd.ExcelFile(file_path)

        # 读取IMP原档工作表，指定没有表头
        df_imp = excel_file.parse('IMP原档', header=None)

        # 检测DataFrame是否为空
        if df_imp.empty:
            raise ValueError("'IMP原档'工作表为空")

        # 设置列名
        df_imp.columns = [chr(65 + i) for i in range(df_imp.shape[1])]

        # 确保A列有数据
        if df_imp['A'].dropna().empty:
            raise ValueError("'IMP原档'工作表的A列全部为空值")

        # 转换A列数据为数值类型，无法转换的会变为NaN
        df_imp['A'] = pd.to_numeric(df_imp['A'], errors='coerce')

        # 丢弃NaN值
        df_imp_clean = df_imp.dropna(subset=['A'])

        if df_imp_clean.empty:
            raise ValueError("'IMP原档'的A列中没有有效的数值")

        # 直接使用二分查找（无需排序，因为数据已排序）
        closest_idx = find_closest_value(df_imp_clean['A'])
        # 获取对应的B列值（注意：现在索引是原始数据的索引）
        corresponding_b_value = df_imp_clean.iloc[closest_idx]['B']

        # 读取ACR工作表
        try:
            df_acr = excel_file.parse('ACR', header=None)
        except ValueError:
            # 如果工作表不存在，创建一个新的DataFrame
            df_acr = pd.DataFrame()

        # 设置或更新A1单元格
        if df_acr.empty:
            df_acr = pd.DataFrame(index=[0], columns=[chr(65 + i) for i in range(26)])

        df_acr.loc[0, 'A'] = corresponding_b_value

        # 写入Excel文件
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_imp.to_excel(writer, sheet_name='IMP原档', header=False, index=False)
            df_acr.to_excel(writer, sheet_name='ACR', header=False, index=False)

        end_time = time.time()  # 结束计时
        print(f"操作完成！耗时: {end_time - start_time:.4f}秒")
        print(f"'IMP原档'A列中最接近600的值对应的B列值为{corresponding_b_value}，已写入'ACR'工作表的A1单元格。")
        return corresponding_b_value

    except Exception as e:
        end_time = time.time()  # 错误发生时结束计时
        print(f"处理Excel文件时出错（耗时: {end_time - start_time:.4f}秒）: {e}")
        return None


# 主程序入口
if __name__ == "__main__":
    file_path = r"E:\System\pic\A报告\IMP数据.xlsx"
    result = process_excel_file(file_path)
    if result is not None:
        print(f"最终结果: {result}")