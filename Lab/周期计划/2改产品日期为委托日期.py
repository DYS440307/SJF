import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import numbers  # 用于设置日期格式


def set_h3_equal_to_g2(folder_path):
    """
    处理指定文件夹及其所有子目录下的Excel文件，将H3单元格的日期设置为与G2单元格相同

    参数:
        folder_path: 根文件夹路径
    """
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"错误: 文件夹 '{folder_path}' 不存在")
        return

    # 遍历所有子文件夹和文件
    for root, _, files in os.walk(folder_path):
        for filename in files:
            # 只处理Excel文件，跳过临时文件
            if filename.endswith(('.xlsx', '.xlsm')) and not filename.startswith('~$'):
                file_path = os.path.join(root, filename)
                relative_path = os.path.relpath(file_path, folder_path)
                print(f"处理文件: {relative_path}")

                try:
                    # 加载Excel文件
                    workbook = load_workbook(file_path)
                    sheet = workbook.active  # 获取第一个工作表

                    # 获取G2单元格的值
                    g2_value = sheet['G2'].value

                    if g2_value:
                        try:
                            # 解析G2单元格的日期
                            if isinstance(g2_value, datetime):
                                # 如果已经是datetime对象，直接使用
                                date_value = g2_value
                            else:
                                # 尝试将字符串转换为datetime对象
                                date_str = str(g2_value).strip()
                                # 尝试常见的日期格式
                                date_formats = ["%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日",
                                                "%m/%d/%Y", "%d-%m-%Y", "%Y%m%d"]
                                date_value = None

                                for fmt in date_formats:
                                    try:
                                        date_value = datetime.strptime(date_str, fmt)
                                        break
                                    except ValueError:
                                        continue

                                if date_value is None:
                                    raise ValueError(f"无法解析日期格式: {date_str}")

                            # 将H3单元格设置为与G2相同的日期
                            sheet['H3'] = date_value
                            # 设置H3单元格为日期格式，保持与G2一致
                            sheet['H3'].number_format = sheet['G2'].number_format

                            print(f"  已将H3单元格设置为与G2相同的日期: {date_value.strftime('%Y-%m-%d')}")

                            # 保存修改
                            workbook.save(file_path)
                        except Exception as e:
                            print(f"  警告: 处理G2单元格日期时出错 - {str(e)}")
                    else:
                        print(f"  警告: G2单元格为空，无法设置H3单元格")

                except InvalidFileException:
                    print(f"  错误: 不是有效的Excel文件")
                except Exception as e:
                    print(f"  处理文件时出错: {str(e)}")


if __name__ == "__main__":
    # 指定要处理的文件夹路径
    folder_path = r"E:\System\desktop\PY\实验室"

    print(f"开始处理文件夹及其子目录: {folder_path}")
    print(f"任务: 将所有Excel文件的H3单元格日期设置为与G2单元格相同\n")
    set_h3_equal_to_g2(folder_path)
    print("\n处理完成")
