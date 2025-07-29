import os
import re
import random
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# --------------------------
# 可动态修改的参数配置
# --------------------------
FOLDER_PATH = r"E:\System\desktop\PY\实验室"  # 要处理的根文件夹路径
K2_CONTENT = "报告编号"  # K2单元格要设置的内容
L3_CONTENT = "EV"  # L3单元格要设置的内容
L2_PREFIX = "SYS"  # L2单元格编号的固定前缀
START_CODE_RANGE = (1, 20)  # 每个子文件夹的起始编码范围 (最小值, 最大值)
SUPPORTED_DATE_FORMATS = [  # 支持的日期格式
    "%Y/%m/%d",
    "%Y-%m-%d",
    "%Y年%m月%d日",
    "%m/%d/%Y"
]
EXCEL_EXTENSIONS = ('.xlsx', '.xlsm')  # 支持的Excel文件扩展名


def process_excel_files():
    """
    处理指定文件夹及其所有子目录下的所有Excel文件，包括：
    1. 修改K2单元格为指定内容
    2. 修改L3单元格为指定内容
    3. 根据G2单元格日期生成L2单元格编号（前缀+年月日+顺序编码）
    4. 处理文件名，在"_"和"."之间插入L2单元格的字段
    """
    # 检查文件夹是否存在
    if not os.path.exists(FOLDER_PATH):
        print(f"错误: 文件夹 '{FOLDER_PATH}' 不存在")
        return

    # 按子文件夹分组处理文件
    subfolder_files = {}

    # 首先收集所有子文件夹及其包含的Excel文件
    for root, _, files in os.walk(FOLDER_PATH):
        excel_files = [f for f in files if f.endswith(EXCEL_EXTENSIONS) and not f.startswith('~$')]
        if excel_files:
            subfolder_files[root] = excel_files

    # 处理每个子文件夹
    for folder, files in subfolder_files.items():
        # 为每个子文件夹生成随机起始编码
        start_code = random.randint(START_CODE_RANGE[0], START_CODE_RANGE[1])
        current_code = start_code

        # 显示当前处理的子文件夹
        rel_path = os.path.relpath(folder, FOLDER_PATH)
        print(f"\n处理子文件夹: {rel_path} (起始编码: {start_code:03d})")

        # 处理当前子文件夹中的所有Excel文件
        for filename in files:
            file_path = os.path.join(folder, filename)
            print(f"处理文件: {filename}")

            # 存储生成的L2值，用于文件名处理
            l2_value = None

            try:
                # 加载Excel文件
                workbook = load_workbook(file_path)
                # 获取第一个工作表
                sheet = workbook.active

                # 修改K2单元格内容
                sheet['K2'] = K2_CONTENT

                # 修改L3单元格内容
                sheet['L3'] = L3_CONTENT

                # 处理L2单元格：根据G2的日期生成编号
                g2_value = sheet['G2'].value
                if g2_value:
                    try:
                        # 尝试解析日期（支持多种格式）
                        if isinstance(g2_value, datetime):
                            date_obj = g2_value
                        else:
                            # 尝试常见日期格式解析
                            date_str = str(g2_value).strip()
                            date_obj = None

                            for fmt in SUPPORTED_DATE_FORMATS:
                                try:
                                    date_obj = datetime.strptime(date_str, fmt)
                                    break
                                except ValueError:
                                    continue

                            if date_obj is None:
                                raise ValueError(f"无法解析日期格式: {date_str}")

                        # 提取年份和月份
                        year = date_obj.year
                        month = date_obj.month

                        # 生成L2单元格的内容
                        l2_value = f"{L2_PREFIX}{year}{month:02d}{current_code:03d}"
                        sheet['L2'] = l2_value
                        print(f"  L2单元格已设置为: {l2_value}")

                        # 递增编码
                        current_code += 1
                    except Exception as e:
                        print(f"  警告: 无法解析G2单元格的日期 - {str(e)}")
                else:
                    print(f"  警告: G2单元格为空，无法生成L2编号")

                # 保存修改
                workbook.save(file_path)
                print(f"  已完成内容修改")

            except InvalidFileException:
                print(f"  错误: 无法处理文件，可能不是有效的Excel文件")
            except Exception as e:
                print(f"  处理文件时出错: {str(e)}")

            # 处理文件名，在"_"和"."之间插入L2单元格的字段
            if l2_value:  # 只有成功生成L2值时才修改文件名
                # 先删除"_"和"."之间原有的内容
                temp_filename = re.sub(r'_.*?\.', '_.', filename)
                # 在"_"和"."之间插入L2值
                new_filename = re.sub(r'_(\.)', f'_{l2_value}\\1', temp_filename)

                if new_filename != filename:
                    old_path = os.path.join(folder, filename)
                    new_path = os.path.join(folder, new_filename)

                    # 检查新文件名是否已存在
                    if os.path.exists(new_path):
                        print(f"  警告: 文件名 '{new_filename}' 已存在，跳过重命名")
                    else:
                        os.rename(old_path, new_path)
                        print(f"  已重命名为: {new_filename}")
                else:
                    print(f"  文件名无需修改")
            else:
                print(f"  未生成L2值，不修改文件名")


if __name__ == "__main__":
    print(f"开始处理文件夹及其子目录: {FOLDER_PATH}")
    process_excel_files()
    print("\n处理完成")
