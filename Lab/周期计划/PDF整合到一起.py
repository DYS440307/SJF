import os
import datetime
from PyPDF2 import PdfMerger
import win32com.client

# ================== 配置 ==================
source_dir = r"E:\System\desktop\PY\实验室\BSA7D70001C0\2026年3月"
output_dir = r"E:\System\desktop\PY\实验室\BSA7D70001C0\2026年3月"
# ========================================


def clean_file_names(folder_path):
    """清理文件名首尾空格"""
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            old_path = os.path.join(root, file)
            new_name = file.strip()

            if file != new_name:
                new_path = os.path.join(root, new_name)

                # 避免重名覆盖
                if not os.path.exists(new_path):
                    os.rename(old_path, new_path)
                    print(f"已修复文件名: {file} -> {new_name}")
                else:
                    print(f"跳过（目标已存在）: {new_name}")


def get_all_files(folder_path):
    """获取所有文件"""
    file_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_list.append(os.path.join(root, file))
    return file_list


def convert_excels_to_pdfs(excel_files):
    """批量Excel转PDF（单Excel进程）"""
    pdf_files = []

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for excel_file in excel_files:
            pdf_path = os.path.splitext(excel_file)[0] + ".pdf"

            try:
                wb = excel.Workbooks.Open(excel_file)

                # 👉 版式优化（防止截断）
                wb.Worksheets.Select()
                excel.ActiveSheet.PageSetup.Zoom = False
                excel.ActiveSheet.PageSetup.FitToPagesWide = 1
                excel.ActiveSheet.PageSetup.FitToPagesTall = False

                wb.ExportAsFixedFormat(0, pdf_path)
                wb.Close(False)

                print(f"转换成功: {pdf_path}")
                pdf_files.append(pdf_path)

            except Exception as e:
                print(f"转换失败: {excel_file} -> {e}")

    finally:
        excel.Quit()

    return pdf_files


def prepare_pdfs(folder_path):
    """准备PDF（必要时执行Excel转PDF）"""
    all_files = get_all_files(folder_path)

    pdf_files = []
    excel_files = []

    for f in all_files:
        if f.lower().endswith(".pdf"):
            pdf_files.append(f)
        elif f.lower().endswith((".xlsx", ".xls")):
            excel_files.append(f)

    # 👉 没有PDF才转换Excel
    if not pdf_files and excel_files:
        print("未检测到PDF，开始Excel转PDF...")
        pdf_files = convert_excels_to_pdfs(excel_files)

    return pdf_files


def get_current_week_number():
    today = datetime.date.today()
    return today.isocalendar()[1]


def merge_pdfs(pdf_files, output_path):
    if not pdf_files:
        print("没有可合并的PDF！")
        return False

    merger = PdfMerger()

    try:
        # 👉 排序（按文件名）
        pdf_files.sort()

        for pdf_file in pdf_files:
            print(f"合并中: {pdf_file}")
            merger.append(pdf_file)

        merger.write(output_path)
        print(f"\n合并完成: {output_path}")
        return True

    except Exception as e:
        print(f"合并失败: {e}")
        return False

    finally:
        merger.close()


if __name__ == "__main__":

    # ✅ 0. 路径检查
    if not os.path.exists(source_dir):
        print(f"源路径不存在: {source_dir}")
        exit()

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # ✅ 1. 清理文件名空格（关键步骤）
    print("开始清理文件名空格...")
    clean_file_names(source_dir)

    # ✅ 2. 准备PDF
    pdf_files_list = prepare_pdfs(source_dir)
    print(f"最终PDF数量: {len(pdf_files_list)}")

    if pdf_files_list:
        # ✅ 3. 输出文件
        current_week = get_current_week_number()
        output_filename = f"KL_{current_week}.pdf"
        output_path = os.path.join(output_dir, output_filename)

        # ✅ 4. 合并
        merge_pdfs(pdf_files_list, output_path)