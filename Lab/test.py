import os
import openpyxl
import secrets
import shutil
from datetime import datetime


def generate_random_numbers(existing_values, value_range):
    min_val, max_val, min_diff = value_range
    max_attempts = 100
    for _ in range(max_attempts):
        # 生成较小值
        smaller_value = round(secrets.SystemRandom().uniform(min_val, max_val - min_diff), 3)
        # 生成较大值，比前者大至少min_diff
        larger_value = round(secrets.SystemRandom().uniform(smaller_value + min_diff, max_val), 3)

        # 检查是否有重复
        if smaller_value not in existing_values and larger_value not in existing_values:
            return smaller_value, larger_value

    raise Exception("无法在100次尝试内生成不重复的随机数")


def process_excel_file(file_path, output_dir):
    try:
        # 打开工作簿
        workbook = openpyxl.load_workbook(file_path)
        # 获取第一个工作表
        sheet = workbook.active

        # 存储已存在的值
        existing_values = set()

        # 定义各范围配置
        RANGE_1 = (130, 142, 2)
        RANGE_2 = (5.8, 6.125, 0.12)
        RANGE_3 = (76.8, 78.6, 0.32)
        RANGE_4 = (9.4, 13.9, 1.21)

        # 对B12:I16的每个单元格对进行操作
        for row in range(12, 17):
            # 生成各列数据
            value_b, value_c = generate_random_numbers(existing_values, RANGE_1)
            value_d, value_e = generate_random_numbers(existing_values, RANGE_2)
            value_g, value_f = generate_random_numbers(existing_values, RANGE_3)
            value_h, value_i = generate_random_numbers(existing_values, RANGE_4)

            # 写入数据
            sheet[f'B{row}'] = value_b
            sheet[f'C{row}'] = value_c
            sheet[f'D{row}'] = value_d
            sheet[f'E{row}'] = value_e
            sheet[f'F{row}'] = value_f
            sheet[f'G{row}'] = value_g
            sheet[f'H{row}'] = value_h
            sheet[f'I{row}'] = value_i

            # 更新已存在的值集合
            existing_values.update([value_b, value_c, value_d, value_e, value_f, value_g, value_h, value_i])

        # 创建输出目录（如果不存在）
        os.makedirs(output_dir, exist_ok=True)

        # 构建输出文件名（添加时间戳）
        file_name = os.path.basename(file_path)
        name_part, ext_part = os.path.splitext(file_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_name = f"{name_part}_{timestamp}{ext_part}"
        output_file_path = os.path.join(output_dir, output_file_name)

        # 保存工作簿到新位置
        workbook.save(output_file_path)
        print(f"成功处理: {file_name} -> {output_file_name}")
        return True
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return False


def main():
    # 源目录
    source_dir = r"Z:\3-品质部\实验室\邓洋枢\1-实验室相关文件\3-周期验证\2025年\小米\S003\模板"
    # 输出目录
    output_dir = r"E:\System\desktop\PY\实验室"

    # 检查源目录是否存在
    if not os.path.exists(source_dir):
        print(f"错误: 源目录不存在 - {source_dir}")
        return

    # 获取所有Excel文件
    excel_files = []
    for root, _, files in os.walk(source_dir):
        for file in files:
            # 修改此处，同时检查"S003"和"模板"
            if (file.lower().endswith(('.xlsx', '.xls')) and
                    "S003" in file and "模板" in file):
                excel_files.append(os.path.join(root, file))

    if not excel_files:
        print("未找到符合条件的Excel文件")
        return

    print(f"找到 {len(excel_files)} 个符合条件的文件")

    # 处理每个Excel文件
    success_count = 0
    for file_path in excel_files:
        if process_excel_file(file_path, output_dir):
            success_count += 1

    print(f"处理完成: 成功 {success_count} 个, 失败 {len(excel_files) - success_count} 个")


if __name__ == "__main__":
    main()