import pandas as pd

# 现在对E:\System\download\物料清单（原档） - 副本_已复制.xlsx中的ABC列处理，从上到下遍历ABC列，举例ABC列中有四行内容完全一样就只保留第一次出现的数据，删除下面的数据，不是删除行
def deduplicate_first_three_columns(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查是否至少有三列
        if df.shape[1] < 3:
            raise ValueError("Excel文件至少需要包含三列数据")

        # 获取前 three 列的列名（将它们视为A、B、C列）
        col_a, col_b, col_c = df.columns[:3]

        # 创建一个集合用于跟踪已出现过的前三列组合
        seen = set()

        # 遍历每一行，检查前三列组合是否已出现
        for index, row in df.iterrows():
            # 获取当前行的前三列值组合
            combo = (row[col_a], row[col_b], row[col_c])

            # 如果组合已出现过，则清空当前行的前三列值
            if combo in seen:
                df.at[index, col_a] = None
                df.at[index, col_b] = None
                df.at[index, col_c] = None
            else:
                seen.add(combo)

        # 保存处理后的文件，添加"_处理后"后缀
        output_path = file_path.replace('.xlsx', '_处理后.xlsx')
        df.to_excel(output_path, index=False)
        print(f"处理完成，文件已保存至: {output_path}")

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")


if __name__ == "__main__":
    # 目标文件路径
    file_path = r"E:\System\download\物料清单（原档） - 副本_已复制.xlsx"
    deduplicate_first_three_columns(file_path)
