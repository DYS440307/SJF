import openpyxl
import os
import unicodedata
import re

# ==============================================================================
# 配置区域 - 以下参数均可根据实际需求修改
# ==============================================================================
# 文件路径配置
# 包含缩写的Excel文件（需要补全的文件）
ABBR_EXCEL_PATH = r"E:\System\desktop\PY\SQE\关系梳理\2_惠州声乐品质履历_IQC检验记录汇总 - 副本.xlsx"
# 包含全称的Excel文件（参考文件）
FULLNAME_EXCEL_PATH = r"E:\System\desktop\PY\SQE\关系梳理\1_采购入库单副本 - 副本.xlsx"

# 列索引配置（A=1, B=2, 以此类推）
ABBR_COLUMN = 2  # 缩写所在列（当前为B列）
FULLNAME_COLUMN = 1  # 全称所在列（当前为A列）
FULLNAME_TARGET_COLUMN = None  # 补全后存放全称的列（None表示替换原缩写列）

# 匹配规则配置
MIN_MATCH_CHARS = 2  # 最少匹配汉字数量（达到此数量才算匹配成功）

# 预处理配置
# 需要替换的文本（键: 要替换的内容, 值: 替换后的内容）
REPLACE_RULES = {
    "韵锦": "昀锦",
    "锦韵": "昀锦"
}

# 需要删除的行规则
DELETE_CONDITIONS = {
    "exact_match": ["科技"],  # 完全匹配这些值的行将被删除
    "contains": ["佳音达"]  # 包含这些值的行将被删除
}

# 调试配置
SHOW_MATCH_SUGGESTIONS = True  # 是否显示每行的匹配结果（True/False）


# ==============================================================================
# 工具函数 - 提供通用功能支持
# ==============================================================================
def clean_text(text):
    """
    文本清理函数：只保留中文字符，去除其他所有字符

    参数:
        text: 原始文本

    返回:
        清理后的文本（仅包含中文字符）
    """
    if text is None:
        return ""

    # 转换为字符串并标准化Unicode字符（处理全角/半角等）
    str_text = str(text)
    normalized = unicodedata.normalize('NFKC', str_text)

    # 提取所有中文字符（正则匹配中文Unicode范围）
    chinese_chars = re.findall(r'[\u4e00-\u9fff]', normalized)

    # 连接成字符串返回
    return ''.join(chinese_chars)


# ==============================================================================
# 预处理函数 - 处理原始数据，为匹配做准备
# ==============================================================================
def preprocess_abbreviation_file(workbook):
    """
    预处理缩写文件：执行文本替换和行删除操作

    参数:
        workbook: 已加载的Excel工作簿对象

    返回:
        预处理后剩余的总行数
    """
    # 获取当前活动工作表
    sheet = workbook.active
    max_row = sheet.max_row  # 原始总行数
    max_col = sheet.max_column  # 总列数

    # --------------------------
    # 步骤1：执行文本替换
    # --------------------------
    print("\n===== 开始预处理：文本替换 =====")
    replace_count = 0  # 替换计数器

    for row in range(1, max_row + 1):
        # 获取当前行缩写列的单元格
        cell = sheet.cell(row=row, column=ABBR_COLUMN)
        original_value = cell.value

        if original_value is not None:
            str_value = str(original_value)
            new_value = str_value  # 初始化为原始值

            # 应用所有替换规则
            for old_str, new_str in REPLACE_RULES.items():
                if old_str in new_value:
                    new_value = new_value.replace(old_str, new_str)
                    replace_count += 1  # 每替换一次计数+1

            # 如果内容有变化，更新单元格值
            if new_value != str_value:
                cell.value = new_value

    print(f"预处理替换完成：共替换 {replace_count} 处")
    # 打印替换规则详情
    for old_str, new_str in REPLACE_RULES.items():
        print(f"  - '{old_str}' → '{new_str}'")

    # --------------------------
    # 步骤2：执行行删除
    # --------------------------
    print("\n===== 开始预处理：删除指定行 =====")
    rows_to_keep = []  # 存储需要保留的行数据

    for row in range(1, max_row + 1):
        # 获取当前行缩写列的值
        cell_value = sheet.cell(row=row, column=ABBR_COLUMN).value
        str_value = str(cell_value) if cell_value is not None else ""

        # 判断是否需要保留此行（不满足任何删除条件则保留）
        keep_row = True

        # 检查完全匹配删除条件
        for keyword in DELETE_CONDITIONS["exact_match"]:
            if str_value.strip() == keyword:
                keep_row = False
                break  # 满足一个条件即可

        # 检查包含匹配删除条件
        if keep_row:  # 如果之前没被标记为删除
            for keyword in DELETE_CONDITIONS["contains"]:
                if keyword in str_value:
                    keep_row = False
                    break

        # 如果需要保留，存储整行数据
        if keep_row:
            row_data = [
                sheet.cell(row=row, column=col).value
                for col in range(1, max_col + 1)
            ]
            rows_to_keep.append(row_data)

    # 计算删除的行数
    deleted_count = max_row - len(rows_to_keep)

    # 清空工作表原有数据
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            sheet.cell(row=row, column=col).value = None

    # 写入保留的行数据
    for new_row, row_data in enumerate(rows_to_keep, start=1):
        for col, value in enumerate(row_data, start=1):
            sheet.cell(row=new_row, column=col).value = value

    # 打印删除结果
    print(f"预处理删除完成：共删除 {deleted_count} 行")
    # 打印删除规则详情
    if DELETE_CONDITIONS["exact_match"]:
        print(f"  - 完全等于以下值的行：{DELETE_CONDITIONS['exact_match']}")
    if DELETE_CONDITIONS["contains"]:
        print(f"  - 包含以下内容的行：{DELETE_CONDITIONS['contains']}")

    return len(rows_to_keep)  # 返回处理后的总行数


# ==============================================================================
# 数据获取函数 - 从Excel中提取所需数据
# ==============================================================================
def get_unique_values(file_path, column_index):
    """
    从指定Excel文件的指定列提取去重后的值列表

    参数:
        file_path: Excel文件路径
        column_index: 列索引（1开始）

    返回:
        去重后的值列表，格式为[(清理后的值, 原始值), ...]
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误: 文件 '{file_path}' 不存在")
        return []

    try:
        # 只读模式打开文件（提高效率）
        workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet = workbook.active

        values = set()  # 使用集合自动去重

        # 遍历指定列的所有行
        for row in sheet.iter_rows(
                min_row=1,
                min_col=column_index,
                max_col=column_index,
                values_only=True
        ):
            value = row[0]  # 行数据是元组，取第一个元素（唯一一列）
            cleaned = clean_text(value)  # 清理文本
            if cleaned:  # 只保留非空值
                values.add((cleaned, value))  # 同时保存清理后的值和原始值

        workbook.close()  # 关闭文件
        return list(values)  # 转换为列表返回

    except Exception as e:
        print(f"读取文件 '{file_path}' 时出错: {str(e)}")
        return []


# ==============================================================================
# 匹配函数 - 实现缩写与全称的匹配逻辑
# ==============================================================================
def count_matching_chars(str1, str2):
    """
    计算两个字符串中相同的汉字数量

    参数:
        str1: 第一个字符串（已清理）
        str2: 第二个字符串（已清理）

    返回:
        相同汉字的数量
    """
    # 计算两个字符串的字符交集
    common_chars = set(str1) & set(str2)
    return len(common_chars)


def find_best_match(abbreviation, fullname_list, min_chars):
    """
    从全称列表中找到与缩写最匹配的项（基于相同汉字数量）

    参数:
        abbreviation: 缩写文本
        fullname_list: 全称列表，格式为[(清理后的值, 原始值), ...]
        min_chars: 最少匹配汉字数量

    返回:
        (最佳匹配的原始全称, 匹配的汉字数量)，无匹配则返回(None, 0)
    """
    # 清理缩写，只保留中文字符
    abbr_clean = clean_text(abbreviation)
    if not abbr_clean:  # 如果缩写为空，直接返回
        return None, 0

    best_match = None
    max_matches = 0  # 记录最大匹配数量

    # 遍历所有全称寻找最佳匹配
    for full_clean, full_original in fullname_list:
        # 计算相同汉字的数量
        match_count = count_matching_chars(abbr_clean, full_clean)

        # 找到匹配数量更多的项
        if match_count >= min_chars and match_count > max_matches:
            max_matches = match_count
            best_match = full_original

    return best_match, max_matches


# ==============================================================================
# 主函数 - 协调各个步骤执行
# ==============================================================================
def complete_abbreviations():
    """主函数：协调执行预处理、数据读取、匹配补全整个流程"""
    try:
        # --------------------------
        # 步骤1：加载文件并执行预处理
        # --------------------------
        print(f"正在打开需要处理的文件: {os.path.basename(ABBR_EXCEL_PATH)}")
        # 加载缩写文件（可写模式）
        workbook = openpyxl.load_workbook(ABBR_EXCEL_PATH, read_only=False, data_only=True)

        # 执行预处理（替换和删除）
        max_row_after_pre = preprocess_abbreviation_file(workbook)
        print(f"预处理完成，剩余 {max_row_after_pre} 行数据\n")

        # --------------------------
        # 步骤2：读取全称列表
        # --------------------------
        print("正在读取全称列表...")
        fullname_list = get_unique_values(FULLNAME_EXCEL_PATH, FULLNAME_COLUMN)
        print(f"从 '{os.path.basename(FULLNAME_EXCEL_PATH)}' 中读取到 {len(fullname_list)} 个独特的全称")

        # 如果没有全称数据，无法进行匹配，直接退出
        if not fullname_list:
            print("没有找到任何全称数据，无法进行匹配")
            workbook.save(ABBR_EXCEL_PATH)
            workbook.close()
            return

        # --------------------------
        # 步骤3：执行缩写补全
        # --------------------------
        sheet = workbook.active  # 获取活动工作表
        max_row = max_row_after_pre  # 预处理后的总行数

        # 统计变量
        completed_count = 0  # 成功补全的数量
        no_match_count = 0  # 未找到匹配的数量
        total_processed = 0  # 总处理数量

        # 确定目标列（补全后的全称存放位置）
        target_column = FULLNAME_TARGET_COLUMN if FULLNAME_TARGET_COLUMN else ABBR_COLUMN

        print("\n开始进行名称匹配补全...")
        # 遍历每一行进行处理
        for row in range(1, max_row + 1):
            # 获取当前行的缩写值
            abbr_cell = sheet.cell(row=row, column=ABBR_COLUMN)
            abbr_value = abbr_cell.value

            # 如果缩写为空，直接标记为未匹配
            if abbr_value is None:
                no_match_count += 1
                continue

            # 查找最佳匹配
            best_match, match_count = find_best_match(
                abbr_value,
                fullname_list,
                MIN_MATCH_CHARS
            )

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

            # 显示进度（每100行显示一次）
            if total_processed % 100 == 0:
                print(f"已处理 {total_processed}/{max_row} 行...")

        # --------------------------
        # 步骤4：保存结果并显示统计信息
        # --------------------------
        workbook.save(ABBR_EXCEL_PATH)  # 保存文件
        workbook.close()  # 关闭文件

        # 显示最终统计结果
        print("\n" + "=" * 50)
        print("所有处理完成！")
        print(f"预处理后总行数: {max_row}")
        print(f"成功补全: {completed_count} 行 (占 {completed_count / total_processed * 100:.1f}%)")
        print(f"未找到匹配: {no_match_count} 行 (占 {no_match_count / total_processed * 100:.1f}%)")
        print(f"结果已保存到: {ABBR_EXCEL_PATH}")

        # 显示结果存放位置
        if FULLNAME_TARGET_COLUMN:
            print(f"补全的全称已保存到第 {chr(64 + target_column)} 列")
        else:
            print("已用全称替换原缩写列（第 {chr(64 + ABBR_COLUMN)} 列）")
        print("=" * 50 + "\n")

    except Exception as e:
        print(f"\n处理错误: {str(e)}")


# ==============================================================================
# 程序入口
# ==============================================================================
if __name__ == "__main__":
    complete_abbreviations()
