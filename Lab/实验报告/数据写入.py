import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 配置参数 - 方便修改
EXCEL_DIR = r"E:\System\pic\A报告"  # Excel文件所在目录
SEARCH_TEXT = "SYS"  # 文件名中要包含的文本
SOURCE_FILE = r"E:\System\pic\A报告\IMP数据.xlsx"  # 源数据文件路径

# 数据映射：源工作表 -> (起始单元格)
SHEET_MAPPING = {
    "Fb": "B12",  # 源文件的Fb工作表数据复制到目标文件第一个工作表的B12起始
    "ACR": "D12",  # 源文件的ACR工作表数据复制到目标文件第一个工作表的D12起始
    "SPL": "F12",  # 源文件的SPL工作表数据复制到目标文件第一个工作表的F12起始
    "THD": "H12"  # 源文件的THD工作表数据复制到目标文件第一个工作表的H12起始
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


def copy_sheet_data(source_wb, source_sheet_name, target_wb, start_cell):
    """将源工作表的全部数据复制到目标工作簿的第一个工作表的指定位置"""
    # 检查源工作表是否存在
    if source_sheet_name not in source_wb:
        print(f"源文件中不存在工作表: {source_sheet_name}")
        return

    # 获取目标工作簿的第一个工作表
    target_ws = target_wb.active

    source_ws = source_wb[source_sheet_name]

    # 解析起始单元格（例如：B12 -> 列索引=2, 行索引=12）
    start_col_letter = start_cell[0]
    start_row = int(start_cell[1:])

    # 将列字母转换为数字索引（A=1, B=2, ...）
    start_col = ord(start_col_letter) - 64  # ASCII值减去64

    # 获取源工作表的最大行和列
    max_row = source_ws.max_row
    max_col = source_ws.max_column

    print(f"从源文件复制工作表 '{source_sheet_name}' 数据到目标文件的第一个工作表，起始位置: {start_cell}")
    print(f"源数据范围: 1-{max_row}行, A-{get_column_letter(max_col)}列")

    # 复制数据
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            # 计算目标单元格位置
            target_row = start_row + row_idx - 1
            target_col = start_col + col_idx - 1

            # 获取源单元格的值
            source_cell = source_ws.cell(row=row_idx, column=col_idx)
            value = source_cell.value

            # 对数值类型的数据保留三位小数
            if isinstance(value, (int, float)):
                # 使用Python的格式化字符串保留三位小数
                value = round(value, 3)
                # 设置Excel单元格的数字格式为三位小数
                target_ws.cell(row=target_row, column=target_col).number_format = '0.000'

            # 获取目标单元格并设置值
            target_ws.cell(row=target_row, column=target_col).value = value

            # 可选：复制单元格样式
            # if source_cell.has_style:
            #     target_ws.cell(row=target_row, column=target_col).font = copy(source_cell.font)
            #     target_ws.cell(row=target_row, column=target_col).border = copy(source_cell.border)
            #     target_ws.cell(row=target_row, column=target_col).fill = copy(source_cell.fill)
            #     target_ws.cell(row=target_row, column=target_col).alignment = copy(source_cell.alignment)

    print(f"成功复制 {max_row} 行 {max_col} 列数据到起始位置 {start_cell}")


def write_to_excel(file_path, cell_data):
    """向Excel文件的指定单元格写入数据，直接覆盖原文件"""
    try:
        # 加载源工作簿
        source_wb = load_workbook(SOURCE_FILE, data_only=True)

        # 加载目标工作簿
        target_wb = load_workbook(file_path)

        # 写入固定数据
        target_ws = target_wb.active
        for cell, value in cell_data.items():
            target_ws[cell] = value

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
                target_ws[f"{next_col_letter}{row}"] = range_str

                # 在控制台打印解析的范围
                print(f"单元格 {cell}: {value} -> 范围: {range_str}")
            else:
                # 处理无法解析的数值类型
                target_ws[f"{next_col_letter}{row}"] = "N/A"
                print(f"单元格 {cell}: {value} -> 范围: N/A (无法解析)")

        # 复制工作表数据
        for source_sheet_name, start_cell in SHEET_MAPPING.items():
            print(f"\n开始复制工作表 '{source_sheet_name}' 到 {start_cell}")
            copy_sheet_data(source_wb, source_sheet_name, target_wb, start_cell)

        # 直接保存并覆盖原文件
        target_wb.save(file_path)
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