import pandas as pd
from datetime import datetime
from openpyxl import load_workbook


def analyze_all_suppliers(data_file_path, score_file_path, target_month=8):
    try:
        # 读取Excel数据文件
        df = pd.read_excel(data_file_path)

        # 列名定义
        date_column = df.columns[0]
        supplier_column = df.columns[1]
        result_column = df.columns[2]

        # 确保日期列被正确解析为日期格式
        df[date_column] = pd.to_datetime(df[date_column])

        # 筛选出目标月份和年份的数据
        target_data = df[(df[date_column].dt.month == target_month) &
                         (df[date_column].dt.year == 2025)]

        # 获取所有不重复的供应商列表
        all_suppliers = target_data[supplier_column].unique()
        print(f"找到的供应商列表：{', '.join(all_suppliers)}")

        # 为每个供应商计算得分
        supplier_scores = {}
        for supplier in all_suppliers:
            # 筛选当前供应商的数据
            supplier_data = target_data[target_data[supplier_column] == supplier]

            # 统计总出现次数
            total_count = len(supplier_data)

            if total_count == 0:
                print(f"{target_month}月中没有找到供应商为'{supplier}'的记录")
                ok_ratio = 0
            else:
                # 统计抽检结果为OK的次数
                ok_count = len(supplier_data[supplier_data[result_column] == "OK"])

                # 计算OK占比
                ok_ratio = ok_count / total_count * 100

                # 输出结果
                print(f"\n{target_month}月中供应商'{supplier}'出现的总次数：{total_count}")
                print(f"其中抽检结果为OK的次数：{ok_count}")
                print(f"抽检结果为OK的占比：{ok_ratio:.2f}%")

            # 计算需要写入的得分（45乘以OK占比）
            score = 45 * (ok_ratio / 100)
            supplier_scores[supplier] = round(score, 2)
            print(f"计算得出的得分：{score:.2f}")

        # 使用openpyxl加载工作簿以保留格式
        wb = load_workbook(score_file_path)

        # 确定目标工作表名称（例如"8月"）
        target_sheet_name = f"{target_month}月"

        # 检查工作表是否存在
        if target_sheet_name not in wb.sheetnames:
            print(f"警告：工作簿中不存在名为'{target_sheet_name}'的工作表")
            return

        # 获取目标工作表
        ws = wb[target_sheet_name]

        # 遍历C列（第3列）查找包含供应商名称的单元格，从第四行开始处理
        for row in range(4, ws.max_row + 1):  # 从第四行开始
            c_cell = ws.cell(row=row, column=3)  # C列
            e_cell = ws.cell(row=row, column=5)  # E列

            if c_cell.value:
                # 检查当前单元格值是否包含任何供应商名称
                cell_value = str(c_cell.value)
                matched = False

                for supplier, score in supplier_scores.items():
                    if supplier in cell_value:
                        e_cell.value = score
                        print(f"已在'{target_sheet_name}'工作表第{row}行写入供应商'{supplier}'的得分：{score}")

                        # 当E列是数值时，设置F、G、H列的值
                        f_cell = ws.cell(row=row, column=6)  # F列是第6列
                        g_cell = ws.cell(row=row, column=7)  # G列是第7列
                        h_cell = ws.cell(row=row, column=8)  # H列是第8列
                        f_cell.value = 35
                        g_cell.value = 20
                        h_cell.value = 0
                        print(f"已在'{target_sheet_name}'工作表第{row}行F列填充35，G列填充20，H列填充0")

                        # 设置I、J、K列的计算公式
                        i_cell = ws.cell(row=row, column=9)  # I列是第9列
                        j_cell = ws.cell(row=row, column=10)  # J列是第10列
                        k_cell = ws.cell(row=row, column=11)  # K列是第11列

                        # 公式中使用当前行号（注意Excel行号从1开始）
                        i_cell.value = f"=100-(E{row}+F{row}+G{row})+H{row}"
                        j_cell.value = f"=100-I{row}"
                        k_cell.value = f'=IF(J{row}<=100,IF(J{row}>=95,"A",IF(J{row}>=80,"B",IF(J{row}>=70,"C",IF(J{row}<70,"D")))),"错误")'

                        print(f"已在'{target_sheet_name}'工作表第{row}行设置I、J、K列计算公式")

                        matched = True
                        break  # 找到匹配的供应商后停止检查其他供应商

                # 如果没有匹配到任何供应商
                if not matched:
                    # E列写入"当月未来料"
                    e_cell.value = "当月未来料"
                    # F到R列（6到18列）也写入"当月未来料"
                    for col in range(6, 19):  # 6是F列，18是R列
                        ws.cell(row=row, column=col).value = "当月未来料"

                    print(f"'{target_sheet_name}'工作表第{row}行C列供应商未匹配，已在E-R列标记为'当月未来料'")

        # 保存修改，保留原有格式
        wb.save(score_file_path)
        print(f"\n已成功将所有得分写入到 {score_file_path} 的'{target_sheet_name}'工作表中，保留了原有格式")

    except Exception as e:
        print(f"处理过程中出现错误：{str(e)}")


if __name__ == "__main__":
    # 数据文件路径
    data_file_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"
    # 评分表文件路径
    score_file_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表 - 副本.xlsx"
    # 目标月份（可以修改为其他月份，如7、9等）
    target_month = 7
    # 调用分析函数
    analyze_all_suppliers(data_file_path, score_file_path, target_month)
