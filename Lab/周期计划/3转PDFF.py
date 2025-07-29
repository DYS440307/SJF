import os
import shutil
import win32com.client
import pythoncom

# --------------------------
# 可动态修改的参数配置
# --------------------------
SOURCE_FOLDER = r"E:\System\desktop\PY\实验室"  # 源Excel文件所在根文件夹
DESTINATION_FOLDER = r"E:\System\desktop\PY\实验室\pdf"  # PDF文件保存根文件夹
EXCEL_EXTENSIONS = ('.xlsx', '.xlsm', '.xls', '.xlsb')  # 支持的Excel文件扩展名
OVERWRITE_EXISTING = False  # 是否覆盖已存在的PDF文件


def ensure_directory_exists(path):
    """确保目录存在，如果不存在则创建"""
    if not os.path.exists(path):
        os.makedirs(path)


def excel_to_pdf(excel_path, pdf_path):
    """
    将单个Excel文件转换为PDF

    参数:
        excel_path: Excel文件路径
        pdf_path: 生成的PDF文件路径
    """
    # 初始化COM对象
    pythoncom.CoInitialize()

    try:
        # 创建Excel应用对象
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # 不显示Excel窗口
        excel.DisplayAlerts = False  # 不显示警告信息

        # 打开Excel文件
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))

        # 转换为PDF
        workbook.ExportAsFixedFormat(
            Type=0,  # 0表示PDF格式
            Filename=os.path.abspath(pdf_path),
            Quality=0  # 0表示标准质量，1表示最小质量
        )

        # 关闭工作簿和Excel
        workbook.Close(SaveChanges=False)
        excel.Quit()

        return True, "转换成功"

    except Exception as e:
        return False, f"转换失败: {str(e)}"

    finally:
        # 释放资源
        pythoncom.CoUninitialize()


def batch_convert_excel_to_pdf():
    """
    批量转换指定文件夹及其子目录中的Excel文件为PDF，
    并按原文件夹结构保存到目标目录
    """
    # 检查源文件夹是否存在
    if not os.path.exists(SOURCE_FOLDER):
        print(f"错误: 源文件夹 '{SOURCE_FOLDER}' 不存在")
        return

    # 确保目标文件夹存在
    ensure_directory_exists(DESTINATION_FOLDER)

    # 统计信息
    total_files = 0
    success_count = 0
    fail_count = 0

    # 遍历所有子文件夹和文件
    for root, _, files in os.walk(SOURCE_FOLDER):
        for filename in files:
            # 检查是否是Excel文件，且不是临时文件
            if filename.lower().endswith(EXCEL_EXTENSIONS) and not filename.startswith('~$'):
                total_files += 1
                excel_path = os.path.join(root, filename)
                relative_path = os.path.relpath(excel_path, SOURCE_FOLDER)

                # 生成对应的PDF保存路径，保持原文件夹结构
                # 获取相对源文件夹的路径
                relative_dir = os.path.relpath(root, SOURCE_FOLDER)
                # 目标文件夹路径 = 根目标文件夹 + 相对路径
                dest_dir = os.path.join(DESTINATION_FOLDER, relative_dir)
                # 确保目标目录存在
                ensure_directory_exists(dest_dir)

                # 生成PDF文件名
                pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                pdf_path = os.path.join(dest_dir, pdf_filename)

                print(f"\n处理文件: {relative_path}")

                # 检查PDF文件是否已存在
                if os.path.exists(pdf_path) and not OVERWRITE_EXISTING:
                    print(f"  跳过: PDF文件已存在于 {os.path.relpath(pdf_path, SOURCE_FOLDER)}")
                    continue

                # 转换为PDF
                success, message = excel_to_pdf(excel_path, pdf_path)

                if success:
                    success_count += 1
                    print(f"  成功: {message} -> {os.path.relpath(pdf_path, SOURCE_FOLDER)}")
                else:
                    fail_count += 1
                    print(f"  失败: {message}")

    # 输出统计结果
    print("\n" + "=" * 50)
    print(f"转换完成 - 总计: {total_files}, 成功: {success_count}, 失败: {fail_count}")
    print(f"所有PDF文件保存在: {DESTINATION_FOLDER}")
    print("=" * 50)


if __name__ == "__main__":
    print(f"开始将Excel文件转换为PDF")
    print(f"源文件夹: {SOURCE_FOLDER}")
    print(f"PDF保存位置: {DESTINATION_FOLDER}")
    print(f"包含子目录: 是")
    print(f"支持的Excel格式: {', '.join(EXCEL_EXTENSIONS)}")
    print(f"覆盖已存在的PDF: {'是' if OVERWRITE_EXISTING else '否'}\n")

    batch_convert_excel_to_pdf()
    print("\n所有文件处理完毕")
