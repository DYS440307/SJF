import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


def process_excel_files():
    # 文件路径
    file1_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表 (2).xlsx"
    file2_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"

    # 检查文件是否存在
    if not os.path.exists(file1_path):
        print(f"错误：文件1不存在 - {file1_path}")
        return
    if not os.path.exists(file2_path):
        print(f"错误：文件2不存在 - {file2_path}")
        return

    try:
        # 加载工作簿以便直接使用单元格引用
        wb1 = load_workbook(file1_path)
        wb2 = load_workbook(file2_path)
        ws1 = wb1.active
        ws2 = wb2.active

        # 获取文件2中B3的供应商名称（B列，第3行）
        supplier_name = ws2['B3'].value
        print(f"要匹配的供应商名称: {supplier_name}")

        if not supplier_name:
            print("错误：文件2的B3单元格为空")
            return

        # 在文件1的C列中查找包含该供应商名称的行
        matched_row = None
        row_index = 1  # 从第1行开始查找

        # 循环查找C列中的内容
        while True:
            cell_value = ws1[f"C{row_index}"].value
            # 如果单元格为空且下一行也为空，则停止查找
            if not cell_value and not ws1[f"C{row_index + 1}"].value:
                break

            # 检查是否包含供应商名称
            if cell_value and supplier_name in str(cell_value):
                matched_row = row_index
                print(f"匹配成功：文件1的第{matched_row}行（C{matched_row}）包含 '{supplier_name}'")
                break

            row_index += 1

        if matched_row is None:
            print(f"未找到匹配项：文件1的C列中没有包含 '{supplier_name}' 的内容")
            wb1.close()
            wb2.close()
            return

        # 获取文件2中J6的值
        j6_value = ws2['J6'].value

        if j6_value is None:
            print("警告：文件2的J6单元格为空")
            j6_value = 0

        # 计算结果
        result = float(j6_value) * 40

        # 写入文件1中匹配行的E列
        target_cell = f"E{matched_row}"
        ws1[target_cell] = result

        # 保存文件1的修改
        wb1.save(file1_path)
        print(f"已将文件2的J6单元格值({j6_value})乘以40后的结果({result})写入文件1的{target_cell}单元格")

        # 关闭工作簿
        wb1.close()
        wb2.close()

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    process_excel_files()
