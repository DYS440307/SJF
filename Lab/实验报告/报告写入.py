import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import os


def find_report_data(report_id, lab_record_file):
    """在老化实验记录中查找报告ID对应的数据行"""
    try:
        print(f"正在读取老化实验记录: {lab_record_file}")
        df = pd.read_excel(lab_record_file)

        # 查找N列中匹配报告ID的行
        report_rows = df[df.iloc[:, 13].astype(str).str.strip() == report_id]

        if report_rows.empty:
            print(f"未找到报告编号为 '{report_id}' 的记录")
            return None

        if len(report_rows) > 1:
            print(f"警告: 找到多条报告编号为 '{report_id}' 的记录，仅使用第一条")

        return report_rows.iloc[0]

    except Exception as e:
        print(f"读取老化实验记录时出错: {e}")
        return None


def parse_i_column(i_value):
    """解析I列数据（格式：TCL；G0202-000313；310100108）"""
    if pd.isna(i_value):
        return [None, None, None]

    parts = str(i_value).split('；')
    parts += [None] * (3 - len(parts))  # 确保返回3个元素
    return parts[:3]


def parse_m_column(m_value):
    """解析M列数据（格式：8.75W；7.24V；直通；振幅冲击1;50°C；90%R.H）"""
    if pd.isna(m_value):
        return [None, None, None, None, None, None]

    # 处理中文和西文分号
    parts = str(m_value).replace(';', '；').split('；')
    parts += [None] * (6 - len(parts))  # 确保返回6个元素
    return parts[:6]


def write_to_report_template(report_data, template_file):
    """将提取的数据写入到试验报告模板中"""
    try:
        os.makedirs(os.path.dirname(template_file), exist_ok=True)
        print(f"正在打开试验报告模板: {template_file}")
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active

        # 基础数据映射
        cell_mapping = {
            'G2': 0,  # A列数据写入G2
            'B4': 1,  # B列数据写入B4
            'D4': 2,  # C列数据写入D4
            'H3': 7,  # H列数据写入H3
            'H4': 9,  # J列数据写入H4
            'J3': 10,  # K列数据写入J3
            'L3': 11,  # L列数据写入L3
            'L2': 13,  # N列数据写入L2
            'B2': 5  # F列数据写入B2（新增）
        }

        # 填充基础数据
        for cell, col_idx in cell_mapping.items():
            if pd.notna(report_data.iloc[col_idx]):
                ws[cell] = report_data.iloc[col_idx]
                print(f"已将数据从列 {get_column_letter(col_idx + 1)} 写入到单元格 {cell}")
            else:
                print(f"列 {get_column_letter(col_idx + 1)} 数据为空，跳过单元格 {cell}")

        # 解析并写入I列数据
        i_parts = parse_i_column(report_data.iloc[8])  # I列索引为8
        i_mapping = {
            'F3': i_parts[0],  # TCL
            'B3': i_parts[1],  # G0202-000313
            'D3': i_parts[2]  # 310100108
        }

        for cell, value in i_mapping.items():
            if value is not None:
                ws[cell] = value
                print(f"已将I列数据 '{value}' 写入到单元格 {cell}")

        # 解析并写入M列数据
        m_parts = parse_m_column(report_data.iloc[12])  # M列索引为12
        m_mapping = {
            'B7': m_parts[0],  # 8.75W
            'C7': m_parts[1],  # 7.24V
            'D7': m_parts[2],  # 直通
            'E7': m_parts[3],  # 振幅冲击1
            'J4': m_parts[4],  # 50°C
            'L4': m_parts[5]  # 90%R.H
        }

        for cell, value in m_mapping.items():
            if value is not None:
                ws[cell] = value
                print(f"已将M列数据 '{value}' 写入到单元格 {cell}")

        # 保存修改后的模板
        output_file = template_file
        wb.save(output_file)
        print(f"试验报告已保存至: {output_file}")

    except Exception as e:
        print(f"写入试验报告时出错: {e}")


def main():
    # 文件路径配置
    LAB_RECORD_FILE = r"Z:\3-品质部\实验室\邓洋枢\2-实验记录汇总表\2025年\老化实验记录.xlsx"
    REPORT_TEMPLATE = r"E:\System\pic\A报告\试验报告.xlsx"

    # 获取用户输入的报告编号
    report_id = input("请输入报告编号: ").strip()
    if not report_id:
        print("报告编号不能为空")
        return

    # 查找报告数据
    report_data = find_report_data(report_id, LAB_RECORD_FILE)
    if report_data is None:
        return

    # 写入报告模板
    write_to_report_template(report_data, REPORT_TEMPLATE)

    print("报告生成完成!")


if __name__ == "__main__":
    main()