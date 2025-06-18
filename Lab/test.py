import os
import openpyxl
import secrets
import shutil
from datetime import datetime


# ===== 配置区域 =====
class Config:
    # 文件路径配置
    SOURCE_DIR = r"Z:\3-品质部\实验室\邓洋枢\1-实验室相关文件\3-周期验证\2025年\小米\S003\模板"
    OUTPUT_DIR = r"E:\System\desktop\PY\实验室"

    # 随机数生成范围配置 (最小值, 最大值, 最小差值)
    RANGE_CONFIG = {
        'B_C': (130, 142, 2),  # B列和C列范围
        'D_E': (5.8, 6.125, 0.12),  # D列和E列范围
        'F_G': (76.8, 78.6, 0.32),  # F列和G列范围
        'H_I': (9.4, 13.9, 1.21),  # H列和I列范围
    }

    # 数据填充区域
    ROW_START = 12
    ROW_END = 16  # 包含此行

    # 文件匹配条件
    FILE_FILTERS = {
        'extensions': ['.xlsx', '.xls'],
        'keywords': ['S003', '模板']
    }


# ===== 功能函数 =====
def generate_random_numbers(existing_values, value_range):
    min_val, max_val, min_diff = value_range
    max_attempts = 100
    for _ in range(max_attempts):
        smaller_value = round(secrets.SystemRandom().uniform(min_val, max_val - min_diff), 3)
        larger_value = round(secrets.SystemRandom().uniform(smaller_value + min_diff, max_val), 3)
        if smaller_value not in existing_values and larger_value not in existing_values:
            return smaller_value, larger_value
    raise Exception("无法在100次尝试内生成不重复的随机数")


def process_excel_file(file_path, output_dir, order_date, order_number, config):
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        sheet = workbook.active

        sheet['G2'] = order_date
        sheet['L2'] = order_number

        existing_values = set()

        for row in range(config.ROW_START, config.ROW_END + 1):
            # 使用配置中的范围生成随机数
            value_b, value_c = generate_random_numbers(existing_values, config.RANGE_CONFIG['B_C'])
            value_d, value_e = generate_random_numbers(existing_values, config.RANGE_CONFIG['D_E'])
            value_g, value_f = generate_random_numbers(existing_values, config.RANGE_CONFIG['F_G'])
            value_h, value_i = generate_random_numbers(existing_values, config.RANGE_CONFIG['H_I'])

            sheet[f'B{row}'] = value_b
            sheet[f'C{row}'] = value_c
            sheet[f'D{row}'] = value_d
            sheet[f'E{row}'] = value_e
            sheet[f'F{row}'] = value_f
            sheet[f'G{row}'] = value_g
            sheet[f'H{row}'] = value_h
            sheet[f'I{row}'] = value_i

            existing_values.update([value_b, value_c, value_d, value_e, value_f, value_g, value_h, value_i])

        os.makedirs(output_dir, exist_ok=True)
        file_name = os.path.basename(file_path)
        new_name = file_name.replace("模板", f"_{order_number}")
        output_file_path = os.path.join(output_dir, new_name)

        workbook.save(output_file_path)
        print(f"成功处理: {file_name} -> {new_name}")
        return True
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return False


def get_input_pairs():
    pairs = []
    print("\n请输入日期和订单编号对（格式：2025/6/12	XSCKD002748）")
    print("每行一对，输入空行结束")

    print("示例输入:")
    print("2025/6/12	XSCKD002748")
    print("2025/6/7	XSCKD002730")
    print("（直接粘贴多行也可以）")

    print("\n开始输入:")
    while True:
        user_input = input().strip()
        if not user_input:
            break

        try:
            parts = user_input.split()
            if len(parts) != 2:
                print("输入格式错误，请使用 '日期 订单编号' 格式")
                continue

            date_str, order_number = parts
            date_parts = date_str.split('/')
            if len(date_parts) == 3:
                year, month, day = date_parts
                formatted_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                pairs.append((formatted_date, order_number))
                print(f"已添加: {formatted_date} {order_number}")
            else:
                print("日期格式错误，请使用 YYYY/MM/DD 格式")
        except Exception as e:
            print(f"输入错误: {e}，请重试")

    return pairs


def get_excel_files(config):
    """根据配置获取符合条件的Excel文件"""
    excel_files = []
    if not os.path.exists(config.SOURCE_DIR):
        print(f"错误: 源目录不存在 - {config.SOURCE_DIR}")
        return excel_files

    for root, _, files in os.walk(config.SOURCE_DIR):
        for file in files:
            # 检查文件扩展名
            if not any(file.lower().endswith(ext) for ext in config.FILE_FILTERS['extensions']):
                continue
            # 检查关键词
            if not all(keyword in file for keyword in config.FILE_FILTERS['keywords']):
                continue
            excel_files.append(os.path.join(root, file))

    return excel_files


def main():
    # 创建配置实例
    config = Config()

    # 移除交互式配置修改
    print(f"\n使用配置:")
    print(f"  源目录: {config.SOURCE_DIR}")
    print(f"  输出目录: {config.OUTPUT_DIR}")

    # 获取输入数据
    input_pairs = get_input_pairs()
    if not input_pairs:
        print("未输入任何数据，程序退出")
        return

    # 获取符合条件的Excel文件
    excel_files = get_excel_files(config)
    if not excel_files:
        print("未找到符合条件的Excel文件")
        return

    print(f"找到 {len(excel_files)} 个符合条件的文件")

    # 批量处理文件
    for order_date, order_number in input_pairs:
        print(f"\n处理订单: {order_date} {order_number}")
        success_count = 0

        for file_path in excel_files:
            if process_excel_file(file_path, config.OUTPUT_DIR, order_date, order_number, config):
                success_count += 1

        print(f"订单 {order_number} 处理完成: 成功 {success_count} 个, 失败 {len(excel_files) - success_count} 个")


if __name__ == "__main__":
    main()