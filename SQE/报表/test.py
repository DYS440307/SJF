import openpyxl
import os


def count_supplier_data(month_sheet, supplier_name):
    """统计指定月份工作表中供应商的总数量和不合格数量"""
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
            if result == "NG":
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

        print(f"正在统计供应商 '{supplier_name}' 的质量数据...")

        # 处理1月份数据
        january_sheet = workbook["1月"]
        jan_total, jan_ng = count_supplier_data(january_sheet, supplier_name)
        trend_sheet["D3"].value = jan_total  # 1月总数量写入D3
        trend_sheet["D4"].value = jan_ng  # 1月不合格数量写入D4
        print(f"1月统计完成：总数量={jan_total}, 不合格数量={jan_ng}")

        # 处理2月份数据
        february_sheet = workbook["2月"]
        feb_total, feb_ng = count_supplier_data(february_sheet, supplier_name)
        trend_sheet["E3"].value = feb_total  # 2月总数量写入E3
        trend_sheet["E4"].value = feb_ng  # 2月不合格数量写入E4
        print(f"2月统计完成：总数量={feb_total}, 不合格数量={feb_ng}")

        # 保存修改
        workbook.save(excel_path)
        print("所有结果已成功写入Excel文件")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    update_supplier_stats()
