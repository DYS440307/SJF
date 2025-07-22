import openpyxl
import os
#抽取12个月份到不良履历

def count_supplier_data(month_sheet, supplier_name):
    """统计指定月份工作表中供应商的总数量和不合格数量，支持供应商名称和NG的模糊匹配"""
    total_count = 0
    ng_count = 0

    # 确保供应商名称为字符串类型，便于后续模糊匹配
    target_supplier = str(supplier_name).lower() if supplier_name else ""

    row = 2  # 从第2行开始，假设第1行是标题
    while True:
        current_supplier = month_sheet[f"B{row}"].value
        result = month_sheet[f"G{row}"].value

        # 如果B列没有值，说明已经到了数据末尾
        if not current_supplier:
            break

        # 对供应商名称进行模糊匹配（不区分大小写）
        current_supplier_str = str(current_supplier).lower()
        if target_supplier in current_supplier_str:
            total_count += 1

            # 模糊匹配NG：将结果转为字符串，不区分大小写检查是否包含"ng"
            if result and "ng" in str(result).lower():
                ng_count += 1

        row += 1

    return total_count, ng_count


def process_supplier(trend_sheet, workbook, supplier_row, months_config):
    """处理单个供应商的数据统计，参数为供应商所在行号"""
    # 供应商名称在B列，当前行
    supplier_name = trend_sheet[f"B{supplier_row}"].value
    if not supplier_name:
        print(f"警告：B{supplier_row}单元格中未找到供应商名称，跳过处理")
        return

    print(f"正在统计包含 '{supplier_name}' 的全年质量数据...")

    # 总数写入当前行，不合格数写入下一行
    total_row = supplier_row
    ng_row = supplier_row + 1

    # 循环处理每个月的数据
    for month_name, total_col, ng_col in months_config:
        # 获取当月工作表
        month_sheet = workbook[month_name]

        # 统计数据
        total, ng = count_supplier_data(month_sheet, supplier_name)

        # 写入结果
        trend_sheet[f"{total_col}{total_row}"].value = total
        trend_sheet[f"{ng_col}{ng_row}"].value = ng

        print(f"{month_name}统计完成：总数量={total}, 不合格数量={ng}")


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

        # 自动处理B列中所有有名称的供应商
        # 从第3行开始检查，直到B列没有数据为止
        current_row = 3
        while True:
            supplier_name = trend_sheet[f"B{current_row}"].value
            if not supplier_name:
                # 如果连续5行都没有供应商名称，则认为已经到了末尾
                empty_count = 0
                temp_row = current_row
                while empty_count < 5 and temp_row < current_row + 5:
                    if not trend_sheet[f"B{temp_row}"].value:
                        empty_count += 1
                    else:
                        break
                    temp_row += 1
                if empty_count >= 5:
                    break

            # 处理当前行的供应商
            process_supplier(trend_sheet, workbook, current_row, months_config)

            # 跳到下一个可能的供应商行（当前供应商行+5，可根据实际情况调整）
            current_row += 5

        # 保存修改
        workbook.save(excel_path)
        print("所有供应商全年数据已成功写入Excel文件")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    update_supplier_stats()
