import pandas as pd
import os


def process_excel(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查必要的列是否存在
        required_columns = ['供应商名称1', '物料编码1', '供应商名称2', '物料编码2']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要的列: {col}")

        # 创建供应商名称2到物料编码2的映射（一个供应商可能对应多个编码）
        supplier2_mapping = {}
        for _, row in df.iterrows():
            supplier = row['供应商名称2']
            code = row['物料编码2']

            if pd.notna(supplier) and pd.notna(code):
                if supplier not in supplier2_mapping:
                    supplier2_mapping[supplier] = set()
                supplier2_mapping[supplier].add(str(code))

        # 处理物料编码1列，合并匹配的物料编码2
        for index, row in df.iterrows():
            supplier1 = row['供应商名称1']

            if pd.notna(supplier1) and supplier1 in supplier2_mapping:
                # 获取当前物料编码1的值
                current_codes = str(row['物料编码1']) if pd.notna(row['物料编码1']) else ""

                # 拆分当前编码为集合（去重）
                current_codes_set = set(current_codes.split(';')) if current_codes else set()

                # 添加匹配的物料编码2
                new_codes = supplier2_mapping[supplier1]
                combined_codes = current_codes_set.union(new_codes)

                # 过滤掉空字符串（如果有的话）
                combined_codes.discard('')

                # 重新组合为字符串
                df.at[index, '物料编码1'] = ';'.join(combined_codes)

        # 生成输出文件路径（在原文件名后加"_processed"）
        dir_name, file_name = os.path.split(file_path)
        base_name, ext = os.path.splitext(file_name)
        output_file = os.path.join(dir_name, f"{base_name}_processed{ext}")

        # 保存处理后的文件
        df.to_excel(output_file, index=False)
        print(f"处理完成，文件已保存至: {output_file}")

        return output_file

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        return None


if __name__ == "__main__":
    # 输入文件路径
    excel_path = r"E:\System\download\1.xlsx"

    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"文件不存在: {excel_path}")
    else:
        process_excel(excel_path)
