import pandas as pd
import numpy as np


def remove_duplicate_groups(file_path, output_path=None):
    """
    处理Excel文件，按A列有值单元格及其下方空白单元格为一组进行去重
    仅根据A列中有数值的单元格内容判断是否重复

    参数:
    file_path: 输入Excel文件路径
    output_path: 输出Excel文件路径，默认为原路径加"_去重后"后缀
    """
    # 如果未指定输出路径，则生成默认路径
    if output_path is None:
        import os
        file_dir, file_name = os.path.split(file_path)
        name, ext = os.path.splitext(file_name)
        output_path = os.path.join(file_dir, f"{name}_去重后{ext}")

    # 读取Excel文件
    try:
        df = pd.read_excel(file_path)
        print(f"成功读取文件: {file_path}")
    except Exception as e:
        print(f"读取文件失败: {e}")
        return

    # 检查文件是否为空
    if df.empty:
        print("Excel文件为空，无需处理")
        return

    # 识别A列有值的行（假设A列是第一列）
    a_column = df.columns[0]
    group_starts = df[a_column].notna()
    group_start_indices = np.where(group_starts)[0]

    if len(group_start_indices) == 0:
        print("A列中没有找到有值的单元格，无需处理")
        return

    # 为每行分配组ID
    group_ids = np.zeros(len(df), dtype=int)
    current_group = 0

    # 为每组分配ID
    for i in range(1, len(group_start_indices)):
        start = group_start_indices[i - 1]
        end = group_start_indices[i]
        group_ids[start:end] = current_group
        current_group += 1

    # 处理最后一组
    group_ids[group_start_indices[-1]:] = current_group

    # 将组ID添加到DataFrame
    df['_group_id'] = group_ids

    # 按组ID分组，仅使用A列中有值的单元格内容作为判断重复的依据
    groups = df.groupby('_group_id')
    group_signatures = {}

    for group_id, group_data in groups:
        # 获取组内A列有值的单元格内容（每个组应该只有一个）
        # 取第一个非空值作为组的标识
        group_identifier = group_data[a_column].dropna().iloc[0]
        group_signatures[group_id] = group_identifier

    # 找出重复的组
    seen_signatures = {}
    duplicate_groups = set()

    for group_id, signature in group_signatures.items():
        if signature in seen_signatures:
            duplicate_groups.add(group_id)
            print(f"发现重复组: 组{group_id} (标识: {signature}) 与 组{seen_signatures[signature]} 重复")
        else:
            seen_signatures[signature] = group_id

    # 保留非重复的组
    df_cleaned = df[~df['_group_id'].isin(duplicate_groups)].drop('_group_id', axis=1)

    # 保存处理后的文件
    try:
        df_cleaned.to_excel(output_path, index=False)
        print(f"去重完成，共删除 {len(duplicate_groups)} 个重复组")
        print(f"处理后的文件已保存至: {output_path}")
    except Exception as e:
        print(f"保存文件失败: {e}")


# 执行去重操作
if __name__ == "__main__":
    input_file = r"E:\System\download\物料清单 - 副本.xlsx"
    remove_duplicate_groups(input_file)
