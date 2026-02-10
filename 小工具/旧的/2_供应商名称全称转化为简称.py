import pandas as pd
import os


def create_supplier_mapping(mapping_file_path):
    """
    从映射文件创建全称到简称的映射字典

    参数:
    mapping_file_path: 供应商名单映射文件路径

    返回:
    supplier_map: 全称到简称的字典
    """
    try:
        # 读取映射文件
        df_mapping = pd.read_excel(mapping_file_path)

        # 确保至少有4列数据
        if len(df_mapping.columns) < 4:
            raise ValueError("映射文件至少需要包含4列（简称、全称、匹配简称、匹配全称）")

        # 获取第三列（匹配简称）和第四列（匹配全称）
        matched_short = df_mapping.iloc[:, 2].fillna('').astype(str)
        matched_full = df_mapping.iloc[:, 3].fillna('').astype(str)

        # 创建映射字典（全称 → 简称）
        supplier_map = {}
        for short, full in zip(matched_short, matched_full):
            # 跳过空值和未匹配的记录
            if full.strip() != '' and full != '未匹配' and short.strip() != '':
                # 确保一个全称只对应一个简称（去重）
                if full not in supplier_map:
                    supplier_map[full] = short

        print(f"✅ 成功创建供应商映射字典，共加载 {len(supplier_map)} 条有效映射关系")
        return supplier_map

    except FileNotFoundError:
        print(f"错误：找不到映射文件 {mapping_file_path}")
        return {}
    except Exception as e:
        print(f"创建映射字典时出错：{str(e)}")
        return {}


def replace_supplier_fullname_with_shortname(purchase_file_path, supplier_map):
    """
    将采购入库单中的供应商全称替换为简称

    参数:
    purchase_file_path: 采购入库单文件路径
    supplier_map: 全称到简称的映射字典
    """
    try:
        # 读取采购入库单文件
        df_purchase = pd.read_excel(purchase_file_path)

        # 检查是否有至少2列数据
        if len(df_purchase.columns) < 2:
            raise ValueError("采购入库单文件至少需要包含2列（物料编码、供应商全称）")

        # 获取第二列（供应商全称）数据
        supplier_fullnames = df_purchase.iloc[:, 1].fillna('').astype(str)

        # 存储替换后的简称
        replaced_shorts = []
        # 存储未找到匹配的记录
        unmatched_records = []

        # 遍历并替换全称
        for idx, full_name in enumerate(supplier_fullnames):
            if full_name.strip() == '':
                replaced_shorts.append('')
                continue

            # 查找对应的简称
            if full_name in supplier_map:
                replaced_shorts.append(supplier_map[full_name])
            else:
                replaced_shorts.append('无对应简称')
                # 记录未匹配的记录
                unmatched_records.append({
                    '行号': idx + 2,  # Excel行号（+2是因为索引从0开始，表头占1行）
                    '供应商全称': full_name
                })

        # 将替换后的简称写入第二列（替换原有全称）
        # 如果想保留原全称，可以写入新列：df_purchase.insert(2, '供应商简称', replaced_shorts)
        df_purchase.iloc[:, 1] = replaced_shorts

        # 保存替换后的文件（可以选择覆盖原文件或保存为新文件）
        # 这里保存为新文件，避免覆盖原文件
        new_file_path = purchase_file_path.replace('.xlsx', '_替换简称后.xlsx')
        df_purchase.to_excel(new_file_path, index=False)

        # 打印处理结果
        print("=" * 60)
        print(f"处理完成！替换后的文件已保存至：{new_file_path}")
        print(f"总共处理了 {len(supplier_fullnames)} 条采购入库记录")

        # 统计替换情况
        matched_count = len(supplier_fullnames) - len(unmatched_records)
        print(f"✅ 成功替换 {matched_count} 条记录的供应商名称")
        print(f"❌ 未找到对应简称 {len(unmatched_records)} 条记录")

        # 打印未匹配的记录
        if len(unmatched_records) > 0:
            print("\n📋 未找到对应简称的记录：")
            print("-" * 40)
            for record in unmatched_records[:20]:  # 只显示前20条，避免输出过长
                print(f"行号：{record['行号']} | 全称：{record['供应商全称']}")
            if len(unmatched_records) > 20:
                print(f"... 还有 {len(unmatched_records) - 20} 条未匹配记录未显示")

        print("=" * 60)

        return df_purchase

    except FileNotFoundError:
        print(f"错误：找不到采购入库单文件 {purchase_file_path}")
    except Exception as e:
        print(f"替换供应商名称时出错：{str(e)}")


# 主程序执行
if __name__ == "__main__":
    # 文件路径配置
    mapping_file = r'E:\System\desktop\供应商名单映射.xlsx'
    purchase_file = r'E:\System\download\采购入库单_2025121111381770_236281.xlsx'

    # 检查文件是否存在
    if not os.path.exists(mapping_file):
        print(f"错误：映射文件不存在 - {mapping_file}")
    elif not os.path.exists(purchase_file):
        print(f"错误：采购入库单文件不存在 - {purchase_file}")
    else:
        # 1. 创建供应商映射字典
        supplier_mapping = create_supplier_mapping(mapping_file)

        if supplier_mapping:
            # 2. 替换采购入库单中的供应商全称
            replace_supplier_fullname_with_shortname(purchase_file, supplier_mapping)