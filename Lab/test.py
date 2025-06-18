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


def process_excel_file(file_path, output_dir, order_date, order_number):
    try:
        # 打开工作簿（保留图片）
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        # 获取第一个工作表
        sheet = workbook.active

        # 写入日期和订单编号
        sheet['G2'] = order_date
        sheet['L2'] = order_number

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

        # 构建输出文件名（用订单编号替换"模板"，不添加时间戳）
        file_name = os.path.basename(file_path)
        new_name = file_name.replace("模板", f"_{order_number}")
        output_file_path = os.path.join(output_dir, new_name)

        # 保存工作簿到新位置
        workbook.save(output_file_path)
        print(f"成功处理: {file_name} -> {new_name}")
        return True
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return False


def get_input_pairs():
    pairs = []
    print("\n请输入日期和订单编号对（格式：2025/6/12 XSCKD002748）")
    print("输入空行结束")

    while True:
        user_input = input("输入日期和订单编号（用空格分隔）: ").strip()
        if not user_input:
            break

        try:
            date_str, order_number = user_input.split(maxsplit=1)
            # 转换日期格式为 YYYY-MM-DD
            parts = date_str.split('/')
            if len(parts) == 3:
                year, month, day = parts
                formatted_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                pairs.append((formatted_date, order_number))
                print(f"已添加: {formatted_date} {order_number}")
            else:
                print("日期格式错误，请使用 YYYY/MM/DD 格式")
        except ValueError:
            print("输入格式错误，请使用 '日期 订单编号' 格式")

    return pairs


def main():
    # 获取用户输入的多对日期和订单编号
    input_pairs = get_input_pairs()

    if not input_pairs:
        print("未输入任何数据，程序退出")
        return

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
            if (file.lower().endswith(('.xlsx', '.xls')) and
                    "S003" in file and "模板" in file):
                excel_files.append(os.path.join(root, file))

    if not excel_files:
        print("未找到符合条件的Excel文件")
        return

    print(f"找到 {len(excel_files)} 个符合条件的文件")

    # 对每对输入执行处理
    for order_date, order_number in input_pairs:
        print(f"\n处理订单: {order_date} {order_number}")
        success_count = 0

        for file_path in excel_files:
            if process_excel_file(file_path, output_dir, order_date, order_number):
                success_count += 1

        print(f"订单 {order_number} 处理完成: 成功 {success_count} 个, 失败 {len(excel_files) - success_count} 个")


if __name__ == "__main__":
    main()