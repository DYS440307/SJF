import os
import re
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


def process_excel_files(folder_path):
    """
    处理指定文件夹及其所有子目录下的所有Excel文件

    参数:
        folder_path: 根文件夹路径
    """
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"错误: 文件夹 '{folder_path}' 不存在")
        return

    # 递归遍历文件夹及其子目录中的所有文件
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            # 检查是否是Excel文件
            if filename.endswith(('.xlsx', '.xlsm')) and not filename.startswith('~$'):
                file_path = os.path.join(root, filename)
                relative_path = os.path.relpath(file_path, folder_path)
                print(f"处理文件: {relative_path}")

                try:
                    # 加载Excel文件
                    workbook = load_workbook(file_path)
                    # 获取第一个工作表
                    sheet = workbook.active

                    # 修改K2单元格为"报告编号"
                    sheet['K2'] = "报告编号"

                    # 修改L3单元格为"EV"
                    sheet['L3'] = "EV"

                    # 保存修改
                    workbook.save(file_path)
                    print(f"  已完成内容修改")

                except InvalidFileException:
                    print(f"  错误: 无法处理文件，可能不是有效的Excel文件")
                except Exception as e:
                    print(f"  处理文件时出错: {str(e)}")

                # 处理文件名，删除"_"和"."之间的字段
                # 使用正则表达式匹配并替换"_"和"."之间的内容
                # 保留"_"，删除中间内容，保留"."及其后面的部分
                new_filename = re.sub(r'_.*?\.', '_.', filename)

                if new_filename != filename:
                    old_path = os.path.join(root, filename)
                    new_path = os.path.join(root, new_filename)

                    # 检查新文件名是否已存在
                    if os.path.exists(new_path):
                        print(f"  警告: 文件名 '{new_filename}' 已存在，跳过重命名")
                    else:
                        os.rename(old_path, new_path)
                        print(f"  已重命名为: {new_filename}")
                else:
                    print(f"  文件名无需修改")


if __name__ == "__main__":
    # 指定要处理的根文件夹路径
    folder_path = r"E:\System\desktop\PY\实验室"

    print(f"开始处理文件夹及其子目录: {folder_path}")
    process_excel_files(folder_path)
    print("处理完成")
