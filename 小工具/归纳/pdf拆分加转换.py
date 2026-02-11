import fitz  # PyMuPDF
from pathlib import Path

# ================== 参数区 ==================
PDF_DIR = Path(r"E:\System\download\新建文件夹")
OUT_DIR = PDF_DIR / "png_output"

DPI = 300  # 300=高清打印级，200=普通高清，400+=超高清
# ===========================================

def pdf_to_png(pdf_path: Path, out_base: Path, dpi: int = 400):
    """
    将 PDF 每一页转换为高质量 PNG
    """
    zoom = dpi / 72  # PyMuPDF 基准 DPI 是 72
    matrix = fitz.Matrix(zoom, zoom)

    doc = fitz.open(pdf_path)

    pdf_name = pdf_path.stem
    pdf_out_dir = out_base / pdf_name
    pdf_out_dir.mkdir(parents=True, exist_ok=True)

    for page_index in range(len(doc)):
        page = doc[page_index]
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        out_png = pdf_out_dir / f"{pdf_name}_page_{page_index + 1}.png"
        pix.save(out_png)

        print(f"已生成: {out_png}")

    doc.close()


def batch_process(pdf_dir: Path, out_dir: Path, dpi: int):
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = list(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        print("未找到 PDF 文件")
        return

    for pdf_file in pdf_files:
        print(f"\n正在处理: {pdf_file}")
        pdf_to_png(pdf_file, out_dir, dpi)

    print("\n全部处理完成")


if __name__ == "__main__":
    batch_process(PDF_DIR, OUT_DIR, DPI)
