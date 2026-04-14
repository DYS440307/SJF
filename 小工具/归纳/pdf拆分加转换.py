import fitz  # PyMuPDF
from pathlib import Path
import multiprocessing as mp
import sys
import os

# ================== 参数区 ==================
DEFAULT_DPI = 300
ROTATE_THRESHOLD = 1.1  # 防误判系数（高度 > 宽度 * 1.1 才旋转）
MAX_WORKERS = max(1, mp.cpu_count() - 1)
# 固定处理的文件夹路径
FIXED_PDF_DIR = r"E:\System\download\飞书下载"
# ===========================================


def pdf_to_png_single(pdf_path: Path, out_base: Path, dpi: int):
    """
    单个 PDF 处理（用于多进程）
    """
    try:
        zoom = dpi / 72
        doc = fitz.open(pdf_path)

        pdf_name = pdf_path.stem
        pdf_out_dir = out_base / pdf_name
        pdf_out_dir.mkdir(parents=True, exist_ok=True)

        for page_index in range(len(doc)):
            page = doc[page_index]

            # ===== 获取页面尺寸 =====
            rect = page.rect
            width = rect.width
            height = rect.height

            # ===== 判断是否旋转 =====
            if height > width * ROTATE_THRESHOLD:
                matrix = fitz.Matrix(zoom, zoom).prerotate(-90)
                rotate_flag = "旋转90°"
            else:
                matrix = fitz.Matrix(zoom, zoom)
                rotate_flag = "未旋转"

            # ===== 渲染 =====
            pix = page.get_pixmap(matrix=matrix, alpha=False)

            out_png = pdf_out_dir / f"{pdf_name}_page_{page_index + 1}.png"
            pix.save(out_png)

        doc.close()
        return f"完成: {pdf_path.name}"

    except Exception as e:
        return f"失败: {pdf_path.name} | 错误: {e}"


def batch_process(pdf_dir: Path, dpi: int):
    """
    多进程批量处理
    """
    if not pdf_dir.exists():
        print("路径不存在")
        return

    out_dir = pdf_dir / "png_output"
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = list(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        print("未找到 PDF 文件")
        return

    print(f"检测到 {len(pdf_files)} 个 PDF，开始处理...")
    print(f"使用进程数: {MAX_WORKERS}\n")

    # ===== 多进程 =====
    with mp.Pool(processes=MAX_WORKERS) as pool:
        results = [
            pool.apply_async(pdf_to_png_single, (pdf, out_dir, dpi))
            for pdf in pdf_files
        ]

        for r in results:
            print(r.get())

    print("\n全部处理完成")


if __name__ == "__main__":
    mp.freeze_support()  # Windows必须

    # 直接使用写死的路径，无需输入
    pdf_dir = Path(FIXED_PDF_DIR)
    print(f"正在处理固定路径：{pdf_dir}\n")
    batch_process(pdf_dir, DEFAULT_DPI)