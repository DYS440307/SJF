import pandas as pd
from datetime import datetime
from openpyxl import load_workbook


def analyze_supplier_data(file_path, score_file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 列名定义
        date_column = df.columns[0]
        supplier_column = df.columns[1]
        result_column = df.columns[2]

        # 确保日期列被正确解析为日期格式
        df[date_column] = pd.to_datetime(df[date_column])

        # 筛选出2025年8月的数据
        august_data = df[(df[date_column].dt.month == 8) &
                         (df[date_column].dt.year == 2025)]

        # 筛选出供应商为"和音"的数据
        heyin_data = august_data[august_data[supplier_column] == "和音"]

        # 统计总出现次数
        total_count = len(heyin_data)

        if total_count == 0:
            print("八月中没有找到供应商为'和音'的记录")
            ok_ratio = 0
        else:
            # 统计抽检结果为OK的次数
            ok_count = len(heyin_data[heyin_data[result_column] == "OK"])

            # 计算OK占比
            ok_ratio = ok_count / total_count * 100

            # 输出结果
            print(f"八月中供应商'和音'出现的总次数：{total_count}")
            print(f"其中抽检结果为OK的次数：{ok_count}")
            print(f"抽检结果为OK的占比：{ok_ratio:.2f}%")

        # 计算需要写入的得分（45乘以OK占比）
        score = 45 * (ok_ratio / 100)
        print(f"计算得出的得分：{score:.2f}")

        # 使用openpyxl加载工作簿以保留格式
        wb = load_workbook(score_file_path)
        ws = wb.active  # 获取活动工作表

        # 遍历C列（第3列）查找包含"和音"的单元格
        # 在对应的E列（第5列）写入得分
        for row in range(1, ws.max_row + 1):
            c_cell = ws.cell(row=row, column=3)  # C列
            if c_cell.value and "和音" in str(c_cell.value):
                e_cell = ws.cell(row=row, column=5)  # E列
                e_cell.value = round(score, 2)
                print(f"已在第{row}行写入得分：{round(score, 2)}")

        # 保存修改，保留原有格式
        wb.save(score_file_path)
        print(f"已成功将得分写入到 {score_file_path} 中，保留了原有格式")

    except Exception as e:
        print(f"处理过程中出现错误：{str(e)}")


if __name__ == "__main__":
    # 数据文件路径
    data_file_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"
    # 评分表文件路径
    score_file_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表 - 副本.xlsx"
    # 调用分析函数
    analyze_supplier_data(data_file_path, score_file_path)
