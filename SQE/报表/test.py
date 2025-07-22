import openpyxl
import os


def count_supplier_data(month_sheet, supplier_name):
    """统计指定月份工作表中供应商的总数量和不合格数量，支持NG模糊匹配"""
    total_count = 0
    ng_count = 0

    row = 2  # 从第2行开始，假设第1行是标题
    while True:
        current_supplier = month_sheet[f"B{row}"].value
        result = month_sheet[f"G{row}"].value

        # 如果B列没有值，说明已经到了数据末尾
        if not current_supplier:
            break

        # 如果是目标供应商，进行统计
        if current_supplier == supplier_name:
            total_count += 1

            # 模糊匹配NG：将结果转为字符串，不区分大小写检查是否包含"ng"
            if result and "ng" in str(result).lower():
                ng_count += 1

        row += 1

    return total_count, ng_count


def update_supplier_stats():
    # Excel文件路径
    excel_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"

    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误：找不到文件 {excel_path}")
        return

    try:
        # 打开Excel文件，使用data_only=False以便能够修改单元格
        workbook = openpyxl.load_workbook(excel_path, data_only=False)

        # 获取需要操作的工作表
        trend_sheet = workbook["供应商质量表现趋势"]

        # 获取B3单元格中的供应商名称
        supplier_name = trend_sheet["B3"].value
        if not supplier_name:
            print("错误：B3单元格中未找到供应商名称")
            return

        print(f"正在统计供应商 '{supplier_name}' 的全年质量数据...")

        # 定义月份对应的工作表名称和目标单元格列
        # 格式：(月份名称, 总数单元格列, 不良数单元格列)
        months_config = [
            ("1月", "D", "D"),
            ("2月", "E", "E"),
            ("3月", "F", "F"),
            ("4月", "G", "G"),
            ("5月", "H", "H"),
            ("6月", "I", "I"),
            ("7月", "J", "J"),
            ("8月", "K", "K"),
            ("9月", "L", "L"),
            ("10月", "M", "M"),
            ("11月", "N", "N"),
            ("12月", "O", "O")
        ]

        # 循环处理每个月的数据
        for month_name, total_col, ng_col in months_config:
            # 获取当月工作表
            month_sheet = workbook[month_name]

            # 统计数据
            total, ng = count_supplier_data(month_sheet, supplier_name)

            # 写入结果（总数写在第3行，不良数写在第4行）
            trend_sheet[f"{total_col}3"].value = total
            trend_sheet[f"{ng_col}4"].value = ng

            print(f"{month_name}统计完成：总数量={total}, 不合格数量={ng}")

        # 保存修改
        workbook.save(excel_path)
        print("全年数据已成功写入Excel文件")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    update_supplier_stats()
