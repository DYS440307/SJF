import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil

# 配置参数 - 方便修改
EXCEL_DIR = r"E:\System\pic\A报告"  # Excel文件所在目录
SEARCH_TEXT = "SYS"  # 文件名中要包含的文本
OUTPUT_DIR = r"E:\System\pic\A报告\修改后"  # 输出目录

# 要写入的数据
CELL_DATA = {
    "B9": "Fc(Hz)",
    "B10": "300±20%",
    "D10": "6±15%",
    "F10": "78±2",
    "H10": "≤10"
}


def find_excel_files(directory, search_text):
    """查找目录中名称包含指定文本的Excel文件"""
    excel_files = []
    for filename in os.listdir(directory):
        if search_text in filename and filename.endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(directory, filename))
    return excel_files


def write_to_excel(input_file_path, output_file_path, cell_data):
    """向Excel文件的指定单元格写入数据，并保存为新文件"""
    try:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

        # 复制原文件到输出路径
        shutil.copy2(input_file_path, output_file_path)

        # 加载工作簿
        wb = load_workbook(output_file_path)
        # 获取第一个工作表
        ws = wb.active

        # 写入数据到指定单元格
        for cell, value in cell_data.items():
            ws[cell] = value

        # 保存工作簿
        wb.save(output_file_path)
        print(f"成功写入数据到 {output_file_path}")
    except Exception as e:
        print(f"处理文件 {input_file_path} 时出错: {e}")


def main():
    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 查找符合条件的Excel文件
    excel_files = find_excel_files(EXCEL_DIR, SEARCH_TEXT)

    if not excel_files:
        print(f"未找到名称包含 '{SEARCH_TEXT}' 的Excel文件")
        return

    print(f"找到 {len(excel_files)} 个Excel文件")

    # 处理每个Excel文件
    for file in excel_files:
        # 获取文件名和扩展名
        filename = os.path.basename(file)
        # 构建输出文件路径
        output_file = os.path.join(OUTPUT_DIR, filename)
        write_to_excel(file, output_file, CELL_DATA)


if __name__ == "__main__":
    main()