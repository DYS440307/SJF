import pandas as pd
import os


def match_supplier_names(file_path):
    """
    匹配供应商简称和全称，将结果写入第三、四列，并打印未匹配的记录

    参数:
    file_path: Excel文件路径
    """
    # 设置pandas显示选项，避免列名截断
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查列数，如果不足4列，添加空列
        while len(df.columns) < 4:
            df[f'第{len(df.columns) + 1}列'] = ''

        # 获取简称和全称列的数据
        short_names = df.iloc[:, 0].fillna('').astype(str)  # 第一列：简称
        full_names = df.iloc[:, 1].fillna('').astype(str)  # 第二列：全称

        # 存储匹配结果
        matched_short = []
        matched_full = []
        # 存储未匹配的记录
        unmatched_records = []

        # 遍历每个简称，查找匹配的全称
        for idx, short in enumerate(short_names):
            if short.strip() == '':
                matched_short.append('')
                matched_full.append('')
                continue

            # 查找包含该简称的全称
            match_found = False
            for full in full_names:
                if short in full and full.strip() != '':
                    matched_short.append(short)
                    matched_full.append(full)
                    match_found = True
                    break

            # 如果没有找到匹配项
            if not match_found:
                matched_short.append(short)
                matched_full.append('未匹配')
                # 记录未匹配的简称及其行号
                unmatched_records.append({
                    '行号': idx + 2,  # Excel行号从1开始，表头占1行，所以+2
                    '简称': short
                })

        # 将匹配结果写入第三列和第四列
        df.iloc[:, 2] = matched_short  # 第三列：匹配的简称
        df.iloc[:, 3] = matched_full  # 第四列：匹配的全称

        # 保存处理后的文件
        df.to_excel(file_path, index=False)

        # 打印处理结果统计
        print("=" * 60)
        print(f"处理完成！文件已保存至：{file_path}")
        print(f"总共处理了 {len(matched_short)} 条记录")

        # 统计匹配情况
        match_count = sum(1 for x in matched_full if x != '未匹配' and x != '')
        unmatched_count = len(unmatched_records)
        print(f"✅ 成功匹配 {match_count} 条记录")
        print(f"❌ 未匹配 {unmatched_count} 条记录")

        # 打印未匹配的记录
        if unmatched_count > 0:
            print("\n📋 未匹配的记录详情：")
            print("-" * 40)
            for record in unmatched_records:
                print(f"行号：{record['行号']} | 简称：{record['简称']}")
        else:
            print("\n🎉 所有记录都已成功匹配！")
        print("=" * 60)

        return df

    except FileNotFoundError:
        print(f"错误：找不到文件 {file_path}")
    except Exception as e:
        print(f"处理过程中出现错误：{str(e)}")


# 主程序执行
if __name__ == "__main__":
    # 文件路径
    file_path = r'E:\System\desktop\供应商名单映射.xlsx'

    # 检查文件是否存在
    if os.path.exists(file_path):
        # 执行匹配处理
        result_df = match_supplier_names(file_path)
    else:
        print(f"错误：文件不存在 - {file_path}")