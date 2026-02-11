import os
import sys
from PyPDF2 import PdfReader, PdfWriter


def split_pdf(source_pdf_path):
    """拆分单个PDF文件的核心函数"""
    # 检查源文件是否存在
    if not os.path.exists(source_pdf_path):
        print(f"❌ 错误：文件 {source_pdf_path} 不存在！")
        return False

    # 提取源文件的目录和文件名（用于创建输出目录）
    source_dir = os.path.dirname(source_pdf_path)
    source_filename = os.path.splitext(os.path.basename(source_pdf_path))[0]
    # 输出目录：源文件同目录下的【文件名_拆分结果】文件夹
    output_dir = os.path.join(source_dir, f"{source_filename}_拆分结果")
    os.makedirs(output_dir, exist_ok=True)

    try:
        # 读取PDF
        reader = PdfReader(source_pdf_path)
        total_pages = len(reader.pages)
        print(f"✅ 开始拆分：{source_filename}.pdf（共{total_pages}页）")

        # 逐页拆分
        for page_num in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[page_num])

            # 拆分后的文件名
            output_filename = f"{source_filename}_第{page_num + 1}页.pdf"
            output_path = os.path.join(output_dir, output_filename)

            with open(output_path, "wb") as f:
                writer.write(f)
            print(f"✅ 已保存：{output_filename}")

        print(f"\n🎉 拆分完成！文件保存在：{output_dir}\n")
        return True

    except Exception as e:
        print(f"❌ 拆分失败：{str(e)}\n")
        return False


def main():
    """主函数：处理拖放的PDF文件（命令行参数）"""
    # 获取命令行参数（拖放的文件路径会作为参数传入）
    args = sys.argv[1:]  # sys.argv[0]是程序自身路径，[1:]是拖放的文件

    if not args:
        # 没有拖放文件时，提示用法
        print("📌 用法：将PDF文件直接拖到本EXE文件上即可自动拆分！")
        print("🔍 提示：支持同时拖放多个PDF文件批量拆分\n")
        os.system("pause")  # 暂停窗口，方便查看提示
        return

    # 遍历所有拖放的文件（支持批量拖放）
    for file_path in args:
        # 只处理PDF文件
        if file_path.lower().endswith(".pdf"):
            split_pdf(file_path)
        else:
            print(f"⚠️ 跳过非PDF文件：{file_path}\n")

    # 拆分完成后暂停窗口，方便查看结果
    os.system("pause")


if __name__ == "__main__":
    main()