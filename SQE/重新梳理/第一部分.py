import openpyxl
import os
import re
import unicodedata

# --------------------------
# 配置参数（在此处修改关键词和规则）
# --------------------------
# 1. 要在A列中匹配并删除的关键词（可修改）
DELETE_A_KEYWORDS = ["劳保"]  # 支持多个关键词，如 ["劳保", "报废"]

# 2. 要在B列中匹配并删除的关键词（新增，可修改）
DELETE_B_KEYWORDS = ["鼓纸胶", "RA溶剂","去渍水","双组份中心胶","双组份内磁磁路胶","天那水","干燥剂","弹波胶","模组"]  # 可添加更多，如 ["鼓纸胶", "RA溶剂", "其他词"]

# 3. 替换规则（可增删改）
REPLACEMENT_RULES = [
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

# 4. 其他配置
MAX_REPLACE_ROUNDS = 3  # 多次替换的最大轮数


def clean_text(text):
    """增强版文本清理函数，移除不可见字符并标准化"""
    if text is None:
        return ""

    str_text = str(text)
    normalized = unicodedata.normalize('NFKC', str_text)  # 处理全角/半角
    cleaned = re.sub(
        r'[\s\u0000-\u001F\u007F-\u009F\u2000-\u200F\u2028-\u202F\u205F-\u206F\uFEFF]',
        '',
        normalized
    )
    return cleaned.lower()


def delete_rows_by_keyword(sheet, column, keywords):
    """
    通用函数：删除指定列中包含任何关键词的行
    参数：
        sheet: 工作表对象
        column: 列名（如"A"、"B"）
        keywords: 关键词列表（如["鼓纸胶", "RA溶剂"]）
    返回：
        删除的行数
    """
    if not keywords:
        return 0

    # 清洗关键词，统一匹配标准
    cleaned_keywords = [clean_text(kw) for kw in keywords]
    max_row = sheet.max_row
    rows_to_delete = []

    # 收集要删除的行号
    for row in range(1, max_row + 1):
        cell_value = sheet[f'{column}{row}'].value
        cell_clean = clean_text(cell_value)

        # 检查是否包含任何关键词
        for kw in cleaned_keywords:
            if kw in cell_clean:
                rows_to_delete.append(row)
                break  # 匹配到一个关键词即可

    # 按行号降序删除（避免索引混乱）
    rows_to_delete.sort(reverse=True)
    deleted_count = 0

    for row in rows_to_delete:
        if row <= sheet.max_row:  # 确保行号有效
            sheet.delete_rows(row)
            deleted_count += 1

    return deleted_count


def process_excel_columns(file_path):
    """完整处理流程：替换→填充→删除→去重→分类"""
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
        print(f"原始数据：共 {max_row} 行，{max_col} 列\n")

        # --------------------------
        # 步骤1：B列字段替换
        # --------------------------
        print("===== 步骤1：B列字段替换 =====")
        cleaned_rules = [(clean_text(k), clean_text(v)) for k, v in REPLACEMENT_RULES if clean_text(k)]
        replaced_count = 0
        upper_shell_count = 0

        # 统一B列为文本格式
        for row in range(1, max_row + 1):
            sheet[f'B{row}'].number_format = '@'

        # 执行替换
        for row in range(1, max_row + 1):
            cell = sheet[f'B{row}']
            original_value = cell.value

            if original_value is not None:
                cleaned_value = clean_text(str(original_value).strip())
                initial_value = cleaned_value

                # 多次替换
                for _ in range(MAX_REPLACE_ROUNDS):
                    changed = False
                    for key, value in cleaned_rules:
                        if key in cleaned_value:
                            cleaned_value = cleaned_value.replace(key, value)
                            replaced_count += 1
                            if key == clean_text("上壳"):
                                upper_shell_count += 1
                            changed = True
                    if not changed:
                        break

                if cleaned_value != initial_value:
                    cell.value = cleaned_value

            if row % 5000 == 0:
                print(f"已处理 {row}/{max_row} 行")

        print(f"替换完成：共替换 {replaced_count} 处，其中'上壳'相关替换 {upper_shell_count} 处\n")

        # --------------------------
        # 步骤2：A列填充
        # --------------------------
        print("===== 步骤2：A列填充 =====")
        current_a_value = None
        for row in range(1, max_row + 1):
            cell_value = sheet[f'A{row}'].value
            if cell_value is not None and str(cell_value).strip() != '':
                current_a_value = cell_value
            elif current_a_value is not None:
                sheet[f'A{row}'].value = current_a_value
        print("A列填充完成\n")

        # --------------------------
        # 步骤3：删除A列包含指定关键词的行
        # --------------------------
        print(f"===== 步骤3：删除A列包含{DELETE_A_KEYWORDS}的行 =====")
        deleted_a = delete_rows_by_keyword(sheet, "A", DELETE_A_KEYWORDS)
        print(f"A列删除完成：共删除 {deleted_a} 行\n")
        max_row = sheet.max_row  # 更新最大行号

        # --------------------------
        # 新增步骤4：删除B列包含指定关键词的行
        # --------------------------
        print(f"===== 步骤4：删除B列包含{DELETE_B_KEYWORDS}的行 =====")
        deleted_b = delete_rows_by_keyword(sheet, "B", DELETE_B_KEYWORDS)
        print(f"B列删除完成：共删除 {deleted_b} 行\n")
        max_row = sheet.max_row  # 再次更新最大行号

        # --------------------------
        # 步骤5：A/B列组合去重
        # --------------------------
        print("===== 步骤5：A/B列组合去重 =====")
        seen_pairs = set()
        duplicate_rows = []

        for row in range(1, max_row + 1):
            a_clean = clean_text(sheet[f'A{row}'].value)
            b_clean = clean_text(sheet[f'B{row}'].value)
            pair = (a_clean, b_clean)

            if not a_clean and not b_clean:
                continue

            if pair in seen_pairs:
                duplicate_rows.append(row)
            else:
                seen_pairs.add(pair)

        # 批量删除重复行
        duplicate_rows.sort(reverse=True)
        batches = []
        if duplicate_rows:
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

        # 执行删除
        deleted_duplicate = 0
        for start_row, length in batches:
            if start_row <= sheet.max_row:
                sheet.delete_rows(start_row, length)
                deleted_duplicate += length

        print(f"去重完成：共删除 {deleted_duplicate} 行重复数据\n")
        max_row = sheet.max_row

        # --------------------------
        # 步骤6：按B列分类
        # --------------------------
        print("===== 步骤6：按B列分类 =====")
        rows_data = []

        for row in range(1, max_row + 1):
            a_val = sheet[f'A{row}'].value
            b_val = sheet[f'B{row}'].value
            if (not a_val or str(a_val).strip() == '') and (not b_val or str(b_val).strip() == ''):
                continue

            row_data = [sheet.cell(row=row, column=col).value for col in range(1, max_col + 1)]
            sort_key = clean_text(b_val) if b_val else chr(127)
            rows_data.append((sort_key, row_data))

        rows_data.sort(key=lambda x: x[0])
        print(f"分类完成：共 {len(rows_data)} 行有效数据")

        # 写入分类后的数据
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                sheet.cell(row=row, column=col).value = None

        for new_row, (_, row_data) in enumerate(rows_data, start=1):
            for col, value in enumerate(row_data, start=1):
                sheet.cell(row=new_row, column=col).value = value

            if new_row % 5000 == 0:
                print(f"已写入 {new_row} 行数据")

        # --------------------------
        # 保存文件
        # --------------------------
        print("\n正在保存文件...")
        workbook.save(file_path)
        print(f"所有处理完成，已覆盖原文件: {file_path}")
        print(
            f"最终统计：删除A列{deleted_a}行，删除B列{deleted_b}行，删除重复{deleted_duplicate}行，剩余{len(rows_data)}行")

    except Exception as e:
        print(f"\n处理错误: {str(e)}")


if __name__ == "__main__":
    EXCEL_PATH = r"E:\System\desktop\PY\SQE\关系梳理\1_采购入库单副本副本.xlsx"  # 可修改文件路径
    process_excel_columns(EXCEL_PATH)
