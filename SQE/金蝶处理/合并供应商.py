import pandas as pd
import os


def fill_excel_column(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 获取A列数据（使用iloc[:, 0]获取第一列，无论列名是什么）
        column_a = df.iloc[:, 0].copy()

        # 记录最后一个有值的单元格内容
        last_value = None

        # 遍历A列，填充数据
        for i in range(len(column_a)):
            current_value = column_a[i]

            # 检查当前单元格是否有值（不是NaN）
            if pd.notna(current_value):
                last_value = current_value
            else:
                # 如果当前单元格为空且存在上次一个有效值，则填充
                if last_value is not None:
                    column_a[i] = last_value

        # 将处理后的A列数据写回DataFrame
        df.iloc[:, 0] = column_a

        # 生成新的文件名，避免覆盖原文件
        directory, filename = os.path.split(file_path)
        name, ext = os.path.splitext(filename)
        new_filename = f"{name}_filled{ext}"
        new_file_path = os.path.join(directory, new_filename)

        # 保存处理后的文件
        df.to_excel(new_file_path, index=False)
        print(f"处理完成，文件已保存至: {new_file_path}")

    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")


if __name__ == "__main__":
    # 目标文件路径
    file_path = r"E:\System\download\采购入库单.xlsx"
    fill_excel_column(file_path)
