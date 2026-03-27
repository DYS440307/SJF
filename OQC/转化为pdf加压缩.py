import os
import sys
import zipfile
from multiprocessing import Pool, cpu_count
import win32com.client
from PIL import Image
from docx2pdf import convert
from reportlab.pdfgen import canvas

# ==============================
# 获取输入路径（支持拖拽）
# ==============================
def get_input_dir():
    if len(sys.argv) > 1:
        path = sys.argv[1].strip('"')
        if os.path.isdir(path):
            return path
        else:
            print("❌ 输入的不是有效文件夹")
            sys.exit()
    else:
        return input("👉 请拖入文件夹路径并回车：").strip('"')

# ==============================
# 自动输出目录
# ==============================
def get_output_dir(input_dir):
    base = os.path.dirname(input_dir)
    output = os.path.join(base, "pdf3")
    os.makedirs(output, exist_ok=True)
    return output

# ==============================
# Excel → PDF（子进程）
# ==============================
def excel_worker(task):
    excel_path, pdf_path = task

    try:
        # 跳过已存在（关键提速点）
        if os.path.exists(pdf_path):
            return pdf_path

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(excel_path)

        for sheet in wb.Worksheets:
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False

        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()

        return pdf_path

    except Exception as e:
        return f"ERROR: {excel_path} | {e}"

# ==============================
# 其他格式转换（主进程）
# ==============================
def image_to_pdf(src, dst):
    if not os.path.exists(dst):
        Image.open(src).convert("RGB").save(dst)
    return dst

def txt_to_pdf(src, dst):
    if os.path.exists(dst):
        return dst

    c = canvas.Canvas(dst)
    with open(src, 'r', encoding='utf-8', errors='ignore') as f:
        y = 800
        for line in f:
            c.drawString(50, y, line.strip())
            y -= 15
            if y < 50:
                c.showPage()
                y = 800
    c.save()
    return dst

def word_to_pdf(src, dst):
    if not os.path.exists(dst):
        convert(src, dst)
    return dst

# ==============================
# 分类文件
# ==============================
def classify_files(folder):
    excel_tasks = []
    other_tasks = []
    pdf_list = []

    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)
        if os.path.isdir(full_path):
            continue

        name, ext = os.path.splitext(file)
        ext = ext.lower()
        pdf_path = os.path.join(folder, name + ".pdf")

        if ext in [".xlsx", ".xls"]:
            excel_tasks.append((full_path, pdf_path))

        elif ext == ".docx":
            other_tasks.append(("word", full_path, pdf_path))

        elif ext in [".jpg", ".jpeg", ".png", ".bmp"]:
            other_tasks.append(("img", full_path, pdf_path))

        elif ext == ".txt":
            other_tasks.append(("txt", full_path, pdf_path))

        elif ext == ".pdf":
            pdf_list.append(full_path)

        else:
            print(f"⚠️ 跳过: {file}")

    return excel_tasks, other_tasks, pdf_list

# ==============================
# 处理非Excel
# ==============================
def process_other(tasks):
    results = []

    for typ, src, dst in tasks:
        try:
            if typ == "word":
                results.append(word_to_pdf(src, dst))
            elif typ == "img":
                results.append(image_to_pdf(src, dst))
            elif typ == "txt":
                results.append(txt_to_pdf(src, dst))
        except Exception as e:
            print(f"❌ 失败: {src} | {e}")

    return results

# ==============================
# 压缩
# ==============================
def zip_files(files, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if isinstance(f, str) and f.endswith(".pdf") and os.path.exists(f):
                z.write(f, os.path.basename(f))

# ==============================
# 主流程
# ==============================
def main():
    print("🚀 极限加速模式启动...")

    INPUT_DIR = get_input_dir()
    OUTPUT_DIR = get_output_dir(INPUT_DIR)

    print(f"📂 输入目录: {INPUT_DIR}")
    print(f"📂 输出目录: {OUTPUT_DIR}")

    excel_tasks, other_tasks, pdf_list = classify_files(INPUT_DIR)

    # ==========================
    # Excel并行（核心加速）
    # ==========================
    process_num = min(4, cpu_count())  # 限制最大4
    print(f"⚡ Excel并发数: {process_num}")

    with Pool(process_num) as p:
        excel_results = p.map(excel_worker, excel_tasks)

    excel_results = [r for r in excel_results if isinstance(r, str) and r.endswith(".pdf")]

    # ==========================
    # 其他文件
    # ==========================
    other_results = process_other(other_tasks)

    # 汇总
    all_pdfs = pdf_list + excel_results + other_results

    if not all_pdfs:
        print("❌ 没有生成PDF")
        return

    # ==========================
    # 用户输入命名
    # ==========================
    month = input("请输入月份（如3）：")
    day = input("请输入日期（如18）：")
    place = input("请输出（如汽车城）")
    zip_name = f"{month}月{day}号（{place}）检验报告.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)

    # ==========================
    # 打包
    # ==========================
    zip_files(all_pdfs, zip_path)

    print(f"✅ 完成！PDF数量: {len(all_pdfs)}")
    print(f"📦 输出文件: {zip_path}")

# ==============================
if __name__ == "__main__":
    main()