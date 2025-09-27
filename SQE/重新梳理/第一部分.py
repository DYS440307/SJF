import openpyxl
import os
import re
import unicodedata


def process_excel_columns(file_path):
    """
    完整处理流程：
    1. B列字段替换（包含深度清洗和格式处理）
    2. A列填充（向下复制直到遇到有内容的单元格）
    3. A/B列组合去重（批量删除重复行）
    4. 按B列内容分类排序
    处理后直接覆盖原文件
    """
    try:
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return

        print("开始加载Excel文件...")
        workbook = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
        sheet = workbook.active
        print("Excel文件加载完成")

        max_row = sheet.max_row
        max_col = sheet.max_column
        print(f"检测到文件共有 {max_row} 行，{max_col} 列数据")

        # --------------------------
        # 1. B列字段替换（核心功能）
        # --------------------------
        print("\n开始进行B列字段替换...")
        replacement_rules = [
            ("上壳端子加工", "上壳"),
            ("L-箱壳组", "箱壳"),
            ("L下壳组", "箱壳"),
            ("R下壳组", "箱壳"),
            ("R-下壳组", "箱壳"),
            ("R-箱壳组", "箱壳"),
            ("上壳组件", "箱壳"),
            ("上壳组", "箱壳"),
            ("下壳组件", "下壳"),
            ("盆架组", "盆架组件"),
            ("箱壳组件", "箱壳"),
            ("上壳", "箱壳"),
            ("下壳", "箱壳"),
            ("减震绵", "减震棉"),
            ("吸音绵", "吸音棉"),
            ("尾数箱", "纸箱"),
            ("外箱", "纸箱"),
            ("鼓纸组件", "鼓纸"),
            ("海绵", "减震棉")
        ]

        replaced_count = 0
        upper_shell_count = 0

        # 先清除B列格式，统一为文本格式
        for row in range(1, max_row + 1):
            cell = sheet[f'B{row}']
            cell.number_format = '@'  # 强制文本格式

        # 执行替换
        for row in range(1, max_row + 1):
            cell = sheet[f'B{row}']
            original_value = cell.value

            if original_value is not None:
                # 深度清洗
                str_original = str(original_value).strip()
                str_normalized = unicodedata.normalize('NFKC', str_original)
                str_clean = re.sub(r'[\s\u200b-\u200f\u2028-\u2029]', '', str_normalized)
                str_clean_lower = str_clean.lower()

                # 尝试匹配替换规则
                matched = False
                for key, value in replacement_rules:
                    key_normalized = unicodedata.normalize('NFKC', key)
                    key_clean = re.sub(r'[\s\u200b-\u200f]', '', key_normalized).lower()

                    if str_clean_lower == key_clean:
                        cell.value = value
                        replaced_count += 1
                        if key == "上壳":
                            upper_shell_count += 1
                        matched = True
                        break

            if row % 1000 == 0:
                print(f"已处理 {row} 行替换...")

        print(f"字段替换完成，共替换 {replaced_count} 处，其中'上壳'替换 {upper_shell_count} 处")

        # --------------------------
        # 2. A列填充
        # --------------------------
        print("\n开始处理A列填充...")
        current_a_value = None
        for row in range(1, max_row + 1):
            cell_value = sheet[f'A{row}'].value
            # 遇到有内容的单元格则更新当前值
            if cell_value is not None and str(cell_value).strip() != '':
                current_a_value = cell_value
            # 空单元格则填充当前值
            elif current_a_value is not None:
                sheet[f'A{row}'].value = current_a_value

        print("A列填充完成")

        # --------------------------
        # 3. A/B列组合去重（批量删除）
        # --------------------------
        print("\n开始处理A/B列组合去重...")
        seen_pairs = set()
        duplicate_rows = []

        for row in range(1, max_row + 1):
            a_val = sheet[f'A{row}'].value
            b_val = sheet[f'B{row}'].value

            # 清洗A/B列值用于去重判断
            a_clean = re.sub(r'[\s\u200b-\u200f]', '', str(a_val).strip()).lower() if a_val is not None else ''
            b_clean = re.sub(r'[\s\u200b-\u200f]', '', str(b_val).strip()).lower() if b_val is not None else ''
            pair = (a_clean, b_clean)

            # 跳过空组合
            if not a_clean and not b_clean:
                continue

            if pair in seen_pairs:
                duplicate_rows.append(row)
            else:
                seen_pairs.add(pair)

        print(f"去重检查完成，发现 {len(duplicate_rows)} 行重复数据")

        # 批量删除重复行（按行号降序）
        if duplicate_rows:
            duplicate_rows.sort(reverse=True)

            # 合并连续行成批次
            batches = []
            current_start = duplicate_rows[0]
            current_length = 1

            for row in duplicate_rows[1:]:
                if row == current_start - 1:
                    current_length += 1
                    current_start = row
                else:
                    batches.append((current_start, current_length))
                    current_start = row
                    current_length = 1
            batches.append((current_start, current_length))

            # 执行批量删除
            deleted_total = 0
            for start_row, length in batches:
                sheet.delete_rows(start_row, length)
                deleted_total += length

            print(f"重复行删除完成，共删除 {deleted_total} 行")
        else:
            print("没有发现重复行，无需删除")

        # --------------------------
        # 4. 按B列内容分类排序
        # --------------------------
        print("\n开始根据B列内容进行分类...")
        max_row = sheet.max_row  # 重新获取删除后的最大行数
        rows_data = []

        for row in range(1, max_row + 1):
            a_val = sheet[f'A{row}'].value
            b_val = sheet[f'B{row}'].value

            # 跳过A和B都为空的行
            if (a_val is None or str(a_val).strip() == '') and (b_val is None or str(b_val).strip() == ''):
                continue

            # 收集整行数据
            row_data = [sheet.cell(row=row, column=col).value for col in range(1, max_col + 1)]
            # 按B列值排序（空值放最后）
            sort_key = str(b_val).strip().lower() if b_val is not None else chr(127)
            rows_data.append((sort_key, row_data))

        # 排序并写入
        rows_data.sort(key=lambda x: x[0])
        print(f"分类完成，共 {len(rows_data)} 行有效数据")

        # 清空工作表并写入分类后的数据
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                sheet.cell(row=row, column=col).value = None

        for new_row, (_, row_data) in enumerate(rows_data, start=1):
            for col, value in enumerate(row_data, start=1):
                sheet.cell(row=new_row, column=col).value = value

            if new_row % 1000 == 0:
                print(f"已写入 {new_row} 行数据...")

        # --------------------------
        # 保存文件（直接覆盖原文件）
        # --------------------------
        print("\n正在保存文件...")
        workbook.save(file_path)
        print(f"所有处理完成，已覆盖原文件: {file_path}")

    except Exception as e:
        print(f"\n处理错误: {str(e)}")


if __name__ == "__main__":
    excel_path = r"E:\System\desktop\PY\SQE\关系梳理\1_采购入库单副本副本.xlsx"
    process_excel_columns(excel_path)
