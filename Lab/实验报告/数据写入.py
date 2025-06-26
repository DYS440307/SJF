import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 配置参数 - 方便修改
EXCEL_DIR = r"E:\System\pic\A报告"  # Excel文件所在目录
SEARCH_TEXT = "SYS"  # 文件名中要包含的文本
SOURCE_FILE = r"E:\System\pic\A报告\IMP数据.xlsx"  # 源数据文件路径

# 数据映射：目标单元格 -> (源工作表, 源单元格)
DATA_MAPPING = {
    "B12": ("Fb", "A1"),  # 从Fb工作表的A1单元格获取数据
    "D12": ("ACR", "A1"),  # 从ACR工作表的A1单元格获取数据
    "F12": ("SPL", "A1"),  # 从SPL工作表的A1单元格获取数据
    "H12": ("THD", "A1")  # 从THD工作表的A1单元格获取数据
}

# 要写入的数据
CELL_DATA = {
    "B9": "Fc(Hz)",
    "B10": "300±20%",
    "D10": "6±15%",
    "F10": "78±2",
    "H10": "≤10"
}


def parse_value_range(value_str):
    """解析值范围字符串，返回最小值和最大值"""
    try:
        if "±" in value_str:
            base, tolerance = value_str.split("±")
            base = float(base)
            if tolerance.endswith('%'):
                tolerance_value = base * float(tolerance.strip('%')) / 100
            else:
                tolerance_value = float(tolerance)
            return base - tolerance_value, base + tolerance_value
        elif "≤" in value_str:
            max_val = float(value_str.replace("≤", ""))
            return 0, max_val
        elif "≥" in value_str:
            min_val = float(value_str.replace("≥", ""))
            return min_val, float('inf')
        else:
            # 默认为精确值
            val = float(value_str)
            return val, val
    except Exception as e:
        print(f"解析值范围失败: {value_str}, 错误: {e}")
        return None, None


def find_excel_files(directory, search_text):
    """查找目录中名称包含指定文本的Excel文件"""
    excel_files = []
    for filename in os.listdir(directory):
        if search_text in filename and filename.endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(directory, filename))
    return excel_files


def get_source_data():
    """从源文件获取数据"""
    try:
        source_data = {}
        wb = load_workbook(SOURCE_FILE, data_only=True)

        for target_cell, (sheet_name, source_cell) in DATA_MAPPING.items():
            if sheet_name in wb:
                ws = wb[sheet_name]
                try:
                    # 直接获取单元格值
                    value = ws[source_cell].value
                    source_data[target_cell] = value
                    print(f"从源文件获取数据: {sheet_name}工作表的{source_cell}单元格 = {value}")
                except Exception as e:
                    print(f"获取{sheet_name}工作表的{source_cell}单元格数据时出错: {e}")
            else:
                print(f"源文件中不存在工作表: {sheet_name}")

        return source_data
    except Exception as e:
        print(f"读取源文件时出错: {e}")
        return {}


def write_to_excel(file_path, cell_data):
    """向Excel文件的指定单元格写入数据，直接覆盖原文件"""
    try:
        # 获取源文件中的数据
        source_data = get_source_data()

        # 加载工作簿
        wb = load_workbook(file_path)
        # 获取第一个工作表
        ws = wb.active

        # 写入数据到指定单元格
        for cell, value in cell_data.items():
            ws[cell] = value

            # 跳过B9单元格的解析
            if cell == "B9":
                continue

            # 计算并写入范围
            row = cell[1:]
            col_letter = cell[0]
            next_col_letter = get_column_letter(ord(col_letter) + 1)

            min_val, max_val = parse_value_range(value)
            if min_val is not None and max_val is not None:
                if min_val == max_val:
                    range_str = f"{min_val}"
                else:
                    range_str = f"{min_val}~{max_val}"
                ws[f"{next_col_letter}{row}"] = range_str

                # 在控制台打印解析的范围
                print(f"单元格 {cell}: {value} -> 范围: {range_str}")
            else:
                # 处理无法解析的数值类型
                ws[f"{next_col_letter}{row}"] = "N/A"
                print(f"单元格 {cell}: {value} -> 范围: N/A (无法解析)")

        # 写入从源文件获取的数据
        for cell, value in source_data.items():
            ws[cell] = value
            print(f"从源文件复制数据到单元格 {cell}: {value}")

        # 直接保存并覆盖原文件
        wb.save(file_path)
        print(f"成功写入数据到 {file_path}")
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")


def main():
    # 查找符合条件的Excel文件
    excel_files = find_excel_files(EXCEL_DIR, SEARCH_TEXT)

    if not excel_files:
        print(f"未找到名称包含 '{SEARCH_TEXT}' 的Excel文件")
        return

    print(f"找到 {len(excel_files)} 个Excel文件")

    # 处理每个Excel文件
    for file in excel_files:
        print(f"\n处理文件: {file}")
        write_to_excel(file, CELL_DATA)


if __name__ == "__main__":
    main()