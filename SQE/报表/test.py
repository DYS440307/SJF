import openpyxl
import os


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
        january_sheet = workbook["1月"]

        # 获取B3单元格中的供应商名称
        supplier_name = trend_sheet["B3"].value
        if not supplier_name:
            print("错误：B3单元格中未找到供应商名称")
            return

        print(f"正在统计供应商 '{supplier_name}' 的质量数据...")

        # 统计1月表中该供应商出现的总次数和不合格(NG)次数
        total_count = 0
        ng_count = 0

        # 遍历1月表中的数据行（假设数据从第2行开始）
        row = 2  # 从第2行开始，因为第1行通常是标题
        while True:
            # 获取B列（供应商名称）和G列（结果）的值
            current_supplier = january_sheet[f"B{row}"].value
            result = january_sheet[f"G{row}"].value

            # 如果B列没有值，说明已经到了数据末尾
            if not current_supplier:
                break

            # 如果是目标供应商，进行统计
            if current_supplier == supplier_name:
                total_count += 1
                if result == "NG":
                    ng_count += 1

            row += 1

        # 将统计结果写入"供应商质量表现趋势"表
        trend_sheet["D3"].value = total_count  # 总数量写入D3
        trend_sheet["D4"].value = ng_count  # 不合格数量写入D4

        # 保存修改
        workbook.save(excel_path)
        print(f"统计完成：总数量={total_count}, 不合格数量={ng_count}")
        print("结果已成功写入Excel文件")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    update_supplier_stats()
