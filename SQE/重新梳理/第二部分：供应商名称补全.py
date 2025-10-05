import openpyxl
import os
import unicodedata
import re

# --------------------------
# 配置参数
# --------------------------
# 包含缩写的Excel文件路径（需要补全的文件）
ABBR_EXCEL_PATH = r"E:\System\desktop\PY\SQE\关系梳理\2_惠州声乐品质履历_IQC检验记录汇总 - 副本.xlsx"
# 包含全称的Excel文件路径（参考文件）
FULLNAME_EXCEL_PATH = r"E:\System\desktop\PY\SQE\关系梳理\1_采购入库单副本 - 副本.xlsx"

# 列索引配置（A=1, B=2, 以此类推）
ABBR_COLUMN = 2  # 缩写所在列（B列）
FULLNAME_COLUMN = 1  # 全称所在列（A列）

# 补全后存放全称的列（如果为None则替换原缩写列）
FULLNAME_TARGET_COLUMN = None  # 例如: 3 表示C列

# 最小匹配字数（至少有多少个相同的汉字才算匹配）
MIN_MATCH_CHARS = 2  # 改为2个字匹配

# 是否显示匹配建议（用于调试）
SHOW_MATCH_SUGGESTIONS = True


def clean_text(text):
    """文本清理函数，只保留中文字符和必要内容"""
    if text is None:
        return ""

    # 转换为字符串并标准化 Unicode 字符
    str_text = str(text)
    normalized = unicodedata.normalize('NFKC', str_text)

    # 只保留中文字符
    chinese_chars = re.findall(r'[\u4e00-\u9fff]', normalized)

    # 连接成字符串返回
    return ''.join(chinese_chars)


def get_unique_values(file_path, column_index):
    """从指定Excel文件的指定列获取去重后的值列表"""
    if not os.path.exists(file_path):
        print(f"错误: 文件 '{file_path}' 不存在")
        return []

    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet = workbook.active

        values = set()  # 使用集合自动去重

        for row in sheet.iter_rows(min_row=1, min_col=column_index, max_col=column_index, values_only=True):
            value = row[0]
            cleaned = clean_text(value)
            if cleaned:  # 只保留非空值
                values.add((cleaned, value))  # 同时保存清理后的值和原始值

        workbook.close()
        # 返回列表: [(清理后的值, 原始值), ...]
        return list(values)

    except Exception as e:
        print(f"读取文件 '{file_path}' 时出错: {str(e)}")
        return []


def count_matching_chars(str1, str2):
    """计算两个字符串中相同的汉字数量"""
    # 计算两个字符串的交集
    common_chars = set(str1) & set(str2)
    return len(common_chars)


def find_best_match(abbreviation, fullname_list, min_chars):
    """寻找至少有min_chars个相同汉字的全称"""
    # 清理缩写，只保留中文字符
    abbr_clean = clean_text(abbreviation)
    if not abbr_clean:
        return None, 0

    best_match = None
    max_matches = 0

    # 遍历所有全称寻找匹配
    for full_clean, full_original in fullname_list:
        # 计算相同汉字的数量
        match_count = count_matching_chars(abbr_clean, full_clean)

        # 找到匹配数量更多的项
        if match_count >= min_chars and match_count > max_matches:
            max_matches = match_count
            best_match = full_original

    return best_match, max_matches


def complete_abbreviations():
    """主函数：匹配并补全缩写"""
    try:
        # 1. 获取所有全称
        print("正在读取全称列表...")
        fullname_list = get_unique_values(FULLNAME_EXCEL_PATH, FULLNAME_COLUMN)
        print(f"从 '{os.path.basename(FULLNAME_EXCEL_PATH)}' 中读取到 {len(fullname_list)} 个独特的全称")

        if not fullname_list:
            print("没有找到任何全称数据，无法进行匹配")
            return

        # 2. 打开需要补全的Excel文件
        print(f"\n正在打开需要补全的文件: {os.path.basename(ABBR_EXCEL_PATH)}")
        workbook = openpyxl.load_workbook(ABBR_EXCEL_PATH, read_only=False, data_only=True)
        sheet = workbook.active

        max_row = sheet.max_row
        print(f"文件包含 {max_row} 行数据，开始处理...")

        # 3. 处理每一行，补全缩写
        completed_count = 0
        no_match_count = 0
        total_processed = 0

        # 确定目标列（如果未指定则使用原缩写列）
        target_column = FULLNAME_TARGET_COLUMN if FULLNAME_TARGET_COLUMN else ABBR_COLUMN

        for row in range(1, max_row + 1):
            # 获取原始缩写
            abbr_cell = sheet.cell(row=row, column=ABBR_COLUMN)
            abbr_value = abbr_cell.value

            if abbr_value is None:
                no_match_count += 1
                continue

            # 查找最佳匹配（至少有MIN_MATCH_CHARS个相同汉字）
            best_match, match_count = find_best_match(abbr_value, fullname_list, MIN_MATCH_CHARS)

            # 显示匹配建议（如果启用）
            if SHOW_MATCH_SUGGESTIONS and match_count > 0:
                print(f"行 {row}: '{abbr_value}' → '{best_match}' (匹配字数: {match_count})")

            # 如果找到匹配项，则补全
            if best_match:
                target_cell = sheet.cell(row=row, column=target_column)
                target_cell.value = best_match
                completed_count += 1

            else:
                if SHOW_MATCH_SUGGESTIONS:
                    print(f"行 {row}: '{abbr_value}' → 未找到足够匹配的项（至少需要{MIN_MATCH_CHARS}个相同汉字）")
                no_match_count += 1

            total_processed += 1

            # 显示进度
            if total_processed % 100 == 0:
                print(f"已处理 {total_processed}/{max_row} 行...")

        # 4. 保存文件
        workbook.save(ABBR_EXCEL_PATH)
        workbook.close()

        # 5. 显示统计结果
        print("\n处理完成！")
        print(f"总处理行数: {total_processed}")
        print(f"成功补全: {completed_count} 行 (占 {completed_count / total_processed * 100:.1f}%)")
        print(f"未找到匹配: {no_match_count} 行 (占 {no_match_count / total_processed * 100:.1f}%)")
        print(f"结果已保存到: {ABBR_EXCEL_PATH}")

        if FULLNAME_TARGET_COLUMN:
            print(f"补全的全称已保存到第 {chr(64 + target_column)} 列")
        else:
            print("已用全称替换原缩写列")

    except Exception as e:
        print(f"\n处理错误: {str(e)}")


if __name__ == "__main__":
    complete_abbreviations()
