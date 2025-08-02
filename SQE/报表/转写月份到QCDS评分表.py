import pandas as pd
import openpyxl


def transfer_supplier_data():
    # 文件路径
    quality_trend_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"
    qcds_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表 (2).xlsx"

    # 工作簿名称
    quality_sheet_name = "供应商质量表现趋势"
    july_sheet_name = "7月"

    try:
        # 读取供应商质量表现趋势数据
        # 使用openpyxl引擎读取，以便后续能准确获取单元格数据
        quality_wb = openpyxl.load_workbook(quality_trend_path, data_only=True)
        quality_sheet = quality_wb[quality_sheet_name]

        # 读取7月工作簿数据
        qcds_wb = openpyxl.load_workbook(qcds_path)
        qcds_sheet = qcds_wb[july_sheet_name]

        # 获取"7月工作簿"中E4单元格对应的供应商名称
        # 注意：openpyxl的单元格索引是从1开始的
        supplier_name = qcds_sheet['E4'].value
        print(f"找到供应商: {supplier_name}")

        # 在"供应商质量表现趋势"工作簿的B列中查找匹配的供应商名称
        # 假设供应商名称在B列，从第2行开始有数据
        found = False
        for row in range(2, quality_sheet.max_row + 1):
            current_supplier = quality_sheet[f'B{row}'].value
            if current_supplier == supplier_name:
                print(f"在第{row}行找到匹配的供应商")

                # 获取J161单元格的值并计算 (这里可能需要根据实际情况调整行号)
                # 注意：用户提到的J161可能是特定数据所在位置，可能需要确认
                j_value = quality_sheet['J161'].value
                if j_value is not None:
                    calculated_value = j_value * 45
                    print(f"计算结果: {j_value} * 45 = {calculated_value}")

                    # 将计算结果填入"7月工作簿"的E4单元格
                    qcds_sheet['E4'].value = calculated_value
                    found = True
                else:
                    print("J161单元格的值为空，无法计算")
                break

        if not found:
            print(f"未找到供应商: {supplier_name}")

        # 保存修改
        qcds_wb.save(qcds_path)
        print("数据已成功更新并保存")

    except Exception as e:
        print(f"发生错误: {str(e)}")
    finally:
        # 确保工作簿被关闭
        if 'quality_wb' in locals():
            quality_wb.close()
        if 'qcds_wb' in locals():
            qcds_wb.close()


if __name__ == "__main__":
    transfer_supplier_data()
