import pandas as pd
import os

# 文件路径
file_path = r"E:\System\download\采购入库单_2025121111381770_236281_替换简称后.xlsx"

# 检查文件是否存在
if not os.path.exists(file_path):
    print(f"错误：文件 {file_path} 不存在！")
else:
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查列数是否足够
        if df.shape[1] < 2:
            print("错误：文件列数不足，需要至少两列（物料编码、供应商）！")
        else:
            # 重命名列（方便处理）
            df.columns = ['物料编码', '供应商'] + list(df.columns[2:])

            # 去除空值行
            df = df.dropna(subset=['物料编码', '供应商'])


            # 按物料编码分组，将供应商用:拼接（去重）
            def merge_suppliers(suppliers):
                # 去重并过滤空字符串
                unique_suppliers = [s for s in suppliers.unique() if str(s).strip() != '']
                return ';'.join(unique_suppliers)


            # 分组聚合
            result_df = df.groupby('物料编码', as_index=False)['供应商'].apply(merge_suppliers)

            # 保存结果到新文件（在原文件名后加_去重后）
            output_path = file_path.replace('.xlsx', '_去重后.xlsx')
            result_df.to_excel(output_path, index=False)

            print(f"处理完成！")
            print(f"原数据行数：{len(df)}")
            print(f"去重后行数：{len(result_df)}")
            print(f"结果已保存至：{output_path}")

    except Exception as e:
        print(f"处理出错：{str(e)}")