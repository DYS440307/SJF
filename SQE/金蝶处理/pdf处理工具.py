import os
import re
import sys
from wand.image import Image
from wand.color import Color

# ==============================
# 文件名规范化函数
# ==============================
def normalize_filename(name):
    base, ext = os.path.splitext(name)
    base = base.replace(" ", "；").replace("_", "；")
    # 开头数字+字母 → 110300006L -> 110300006-L
    base = re.sub(r'^(\d+)([A-Za-z])', r'\1-\2', base)
    # 只处理分开的 A；0 -> A0
    base = re.sub(r'；([A-Za-z])；0；', r'；\10；', base)
    # 数字/字母/中文间加分隔符
    base = re.sub(r'(?<=[0-9])(?=[\u4e00-\u9fff])', '；', base)
    base = re.sub(r'(?<=[A-Za-z])(?=[\u4e00-\u9fff])', '；', base)
    base = re.sub(r'(?<=[\u4e00-\u9fff])(?=[A-Za-z0-9])', '；', base)
    base = re.sub(r'；{2,}', '；', base)
    return base + ext

# ==============================
# PDF 转长图函数
# ==============================
def process_single_pdf(pdf_path, dpi=800):
    try:
        file_dir, file_name = os.path.split(pdf_path)
        base_name = os.path.splitext(file_name)[0]
        img_path = os.path.join(file_dir, f"{base_name}.png")

        pages_images = []

        # 打开 PDF，每页处理
        with Image(filename=pdf_path, resolution=dpi) as pdf:
            for i, page in enumerate(pdf.sequence):
                with Image(page) as img:
                    img.background_color = Color("white")
                    img.alpha_channel = 'remove'
                    img.trim()
                    pages_images.append(img.clone())

        total_height = sum(img.height for img in pages_images)
        max_width = max(img.width for img in pages_images)

        # 拼接为长图
        with Image(width=max_width, height=total_height, background=Color("white")) as final_img:
            y_offset = 0
            for img in pages_images:
                final_img.composite(img, left=0, top=y_offset)
                y_offset += img.height
            final_img.save(filename=img_path)
            print(f"✅ 生成长图: {img_path}")

        # 删除原 PDF
        os.remove(pdf_path)
        print(f"🗑 删除原PDF: {pdf_path}\n")

    except Exception as e:
        print(f"❌ 处理 {pdf_path} 时出错: {str(e)}\n")

# ==============================
# 批量处理文件夹内 PDF
# ==============================
def process_all_pdfs_in_folder(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file_name)
            process_single_pdf(pdf_path)

# ==============================
# 批量规范化文件名
# ==============================
def normalize_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            old_path = os.path.join(root, file)
            new_name = normalize_filename(file)
            if new_name != file:
                new_path = os.path.join(root, new_name)
                try:
                    os.rename(old_path, new_path)
                    print(f"✅ 重命名: {file} → {new_name}")
                except Exception as e:
                    print(f"❌ {file} 重命名失败: {e}")

# ==============================
# 文件夹一体化处理
# ==============================
def process_folder(folder_path):
    print(f"开始处理文件夹: {folder_path}")
    process_all_pdfs_in_folder(folder_path)  # 先处理 PDF
    normalize_folder(folder_path)            # 再规范化文件名

# ==============================
# 主程序入口
# ==============================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("请拖拽文件或文件夹到本程序上运行。")
        input("按回车键退出...")
        sys.exit()

    target = sys.argv[1]

    if os.path.isfile(target) and target.lower().endswith(".pdf"):
        print(f"开始处理单个PDF: {target}")
        process_single_pdf(target)
        # 文件名规范化
        dir_path = os.path.dirname(target)
        normalize_folder(dir_path)
    elif os.path.isdir(target):
        process_folder(target)
    else:
        print("输入既不是PDF文件，也不是文件夹。")

    input("\n处理完成，按回车退出...")
