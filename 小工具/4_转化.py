import pandas as pd
import os

# 文件路径
file_path = r"E:\System\download\采购入库单_2025121111381770_236281_替换简称后_去重后.xlsx"

# 检查文件是否存在
if not os.path.exists(file_path):
    print(f"错误：文件不存在 - {file_path}")
    exit(1)

try:
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 查看前几行数据，确认列结构
    print("原始数据前5行：")
    print(df.head())
    print("\n列名：", df.columns.tolist())

    # 1. 创建物料编码到供应商的映射字典（1列=物料编码，2列=供应商）
    # 去除空值行
    mapping_df = df[[df.columns[0], df.columns[1]]].dropna(subset=[df.columns[0]])
    # 创建映射字典（如果有重复的物料编码，取最后一个对应的供应商）
    material_supplier_map = dict(zip(mapping_df[df.columns[0]], mapping_df[df.columns[1]]))

    print(f"\n共创建 {len(material_supplier_map)} 个物料编码-供应商映射关系")

    # 2. 根据第三列的物料编码匹配供应商，填充到第四列
    # 检查第三列是否存在
    if len(df.columns) < 3:
        print("错误：文件至少需要3列数据")
        exit(1)

    # 创建第四列（如果不存在则新增）
    if len(df.columns) >= 4:
        fourth_col_name = df.columns[3]
    else:
        fourth_col_name = "匹配的供应商"
        df[fourth_col_name] = ""

    # 填充第四列：根据第三列的物料编码匹配供应商，未匹配则留空
    df[fourth_col_name] = df[df.columns[2]].map(material_supplier_map).fillna("")

    # 3. 保存处理后的文件（在原文件名后添加_已处理）
    file_dir, file_name = os.path.split(file_path)
    file_name_parts = file_name.rsplit(".", 1)
    new_file_name = f"{file_name_parts[0]}_已处理.{file_name_parts[1]}"
    new_file_path = os.path.join(file_dir, new_file_name)

    # 保存Excel文件
    df.to_excel(new_file_path, index=False)

    print(f"\n处理完成！")
    print(f"原始文件：{file_path}")
    print(f"处理后文件：{new_file_path}")

    # 显示处理结果统计
    matched_count = df[fourth_col_name].notna().sum() - df[fourth_col_name].eq("").sum()
    total_count = df[df.columns[2]].notna().sum()
    print(f"\n匹配统计：")
    print(f"第三列有效物料编码数：{total_count}")
    print(f"成功匹配供应商数：{matched_count}")
    print(f"未匹配数：{total_count - matched_count}")

except Exception as e:
    print(f"处理过程中出错：{str(e)}")