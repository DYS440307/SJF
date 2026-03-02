import os
import datetime
from PyPDF2 import PdfMerger

# 配置参数
source_dir = r"E:\System\desktop\PY\实验室\PDF输出"  # 源PDF文件夹路径
output_dir = r"E:\System\desktop\PY\实验室"  # 输出文件夹路径


def get_all_pdf_files(folder_path):
    """遍历文件夹及其子文件夹，获取所有PDF文件路径"""
    pdf_files = []
    # 遍历目录树
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 筛选出PDF文件
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files


def get_current_week_number():
    """获取当前日期的ISO周数（1-53）"""
    today = datetime.date.today()
    # isocalendar() 返回 (年, 周数, 星期几)
    week_number = today.isocalendar()[1]
    return week_number


def merge_pdfs(pdf_files, output_path):
    """合并多个PDF文件"""
    if not pdf_files:
        print("未找到任何PDF文件！")
        return False

    merger = PdfMerger()
    try:
        # 逐个添加PDF文件
        for pdf_file in pdf_files:
            print(f"正在合并: {pdf_file}")
            merger.append(pdf_file)

        # 写入合并后的文件
        merger.write(output_path)
        print(f"\nPDF合并完成！文件保存至: {output_path}")
        return True
    except Exception as e:
        print(f"合并PDF时出错: {e}")
        return False
    finally:
        # 关闭合并器
        merger.close()


if __name__ == "__main__":
    # 1. 获取所有PDF文件
    pdf_files_list = get_all_pdf_files(source_dir)
    print(f"共找到 {len(pdf_files_list)} 个PDF文件")

    if pdf_files_list:
        # 2. 获取当前周数
        current_week = get_current_week_number()

        # 3. 构建输出文件名和路径
        output_filename = f"KL_{current_week}.pdf"
        output_path = os.path.join(output_dir, output_filename)

        # 4. 合并PDF
        merge_pdfs(pdf_files_list, output_path)