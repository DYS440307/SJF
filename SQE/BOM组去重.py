import pandas as pd
import numpy as np


def remove_duplicate_groups(input_file, output_file, sheet_name=0):
    # 读取数据（指定列索引0、1、2，对应A、B、C列，无表头）
    df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=[0, 1, 2], header=None)
    # 给列命名（方便后续处理，0对应A列，1对应B列，2对应C列）
    df.columns = ['A', 'B', 'C']

    # 处理空值（将A列空值转为NaN便于判断）
    df['A'] = df['A'].replace('', np.nan)

    # 识别组的起始位置（A列非空的行索引）
    group_starts = df.index[df['A'].notna()].tolist()
    if not group_starts:
        print("未找到任何组（A列无有效值）")
        return

    # 确定每组的行范围
    group_starts.append(len(df))
    groups = []

    for i in range(len(group_starts) - 1):
        start = group_starts[i]
        end = group_starts[i + 1] - 1
        group_data = df.loc[start:end, ['B', 'C']]
        group_feature = tuple(tuple(row) for row in group_data.values)
        groups.append({
            'start': start,
            'end': end,
            'feature': group_feature,
            'data': group_data
        })

    # 筛选唯一组
    seen_features = set()
    unique_groups = []
    for group in groups:
        if group['feature'] not in seen_features:
            seen_features.add(group['feature'])
            unique_groups.append(group)

    # 合并结果并输出
    all_rows = []
    for group in unique_groups:
        all_rows.extend(range(group['start'], group['end'] + 1))
    result_df = df.loc[all_rows].reset_index(drop=True)
    # 输出时不保留临时列名（如果需要）
    result_df.to_excel(output_file, index=False, header=None)
    print(f"处理完成！共识别{len(groups)}组，去重后保留{len(unique_groups)}组，结果已保存至{output_file}")


if __name__ == "__main__":
    input_file = r"E:\System\download\组装BOM五表头.xlsx"  # 注意路径前加r避免转义问题
    output_file = r"E:\System\download\组装BOM五表头2.xlsx"
    remove_duplicate_groups(input_file, output_file)