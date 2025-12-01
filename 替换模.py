from PyPDF2 import PdfReader, PdfWriter
import os


def get_pdf_page_count(pdf_path):
    """获取PDF文件的总页数"""
    try:
        # 检查文件是否存在
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"文件不存在: {pdf_path}")

        # 打开PDF并获取页数
        reader = PdfReader(pdf_path)
        page_count = len(reader.pages)
        print(f"\n✅ PDF文件 '{os.path.basename(pdf_path)}' 的总页数: {page_count} 页")
        return page_count
    except Exception as e:
        print(f"❌ 获取PDF页数失败: {str(e)}")
        return None


def replace_pdf_pages(original_pdf_path, replacement_pdf_path, start_page, end_page, output_pdf_path):
    """
    用替换PDF替代原始PDF中指定的页数范围

    参数说明：
    original_pdf_path: 原始PDF文件路径
    replacement_pdf_path: 用于替换的PDF文件路径
    start_page: 替换起始页（从1开始计数）
    end_page: 替换结束页（包含该页）
    output_pdf_path: 新生成的PDF文件路径
    """
    try:
        # 验证文件是否存在
        if not os.path.exists(original_pdf_path):
            raise FileNotFoundError(f"原始PDF文件不存在: {original_pdf_path}")
        if not os.path.exists(replacement_pdf_path):
            raise FileNotFoundError(f"替换PDF文件不存在: {replacement_pdf_path}")

        # 获取原始PDF和替换PDF的页数
        original_reader = PdfReader(original_pdf_path)
        replacement_reader = PdfReader(replacement_pdf_path)

        original_page_count = len(original_reader.pages)
        replacement_page_count = len(replacement_reader.pages)

        # 验证页码范围有效性
        if start_page < 1 or end_page < start_page:
            raise ValueError(f"无效的页码范围！起始页({start_page})必须≥1，且结束页({end_page})≥起始页")

        if end_page > original_page_count:
            raise ValueError(f"结束页({end_page})超过原始PDF总页数({original_page_count})")

        # 验证替换PDF的页数是否足够
        required_pages = end_page - start_page + 1
        if replacement_page_count < required_pages:
            raise ValueError(
                f"替换PDF页数不足！需要替换{required_pages}页，但替换PDF只有{replacement_page_count}页"
            )

        # 创建PDF写入器
        writer = PdfWriter()

        # 1. 添加原始PDF中替换范围之前的页面（1~start_page-1）
        for page_num in range(start_page - 1):
            writer.add_page(original_reader.pages[page_num])

        # 2. 添加替换PDF的页面（按需要的页数）
        for page_num in range(required_pages):
            writer.add_page(replacement_reader.pages[page_num])

        # 3. 添加原始PDF中替换范围之后的页面（end_page~末尾）
        for page_num in range(end_page, original_page_count):
            writer.add_page(original_reader.pages[page_num])

        # 保存新PDF文件
        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)

        print(f"\n✅ 操作成功！新PDF已保存至: {output_pdf_path}")
        print(f"📋 操作详情：")
        print(f"   - 原始PDF：{os.path.basename(original_pdf_path)}（{original_page_count}页）")
        print(f"   - 替换PDF：{os.path.basename(replacement_pdf_path)}（{replacement_page_count}页）")
        print(f"   - 替换范围：第{start_page}页 ~ 第{end_page}页（共{required_pages}页）")

    except Exception as e:
        print(f"\n❌ 替换PDF页面失败: {str(e)}")


def main():
    print("=" * 60)
    print("                PDF页数识别与页面替换工具")
    print("=" * 60)

    # 1. 输入文件路径
    original_pdf = input("\n请输入原始PDF文件路径（可拖入文件）：").strip().replace('"', '')
    replacement_pdf = input("请输入替换PDF文件路径（可拖入文件）：").strip().replace('"', '')

    # 2. 获取原始PDF页数
    original_page_count = get_pdf_page_count(original_pdf)
    if not original_page_count:
        return

    # 3. 输入替换页码范围
    while True:
        try:
            start_page = int(input("\n请输入替换起始页（从1开始）：").strip())
            end_page = int(input("请输入替换结束页（包含该页）：").strip())
            break
        except ValueError:
            print("❌ 请输入有效的数字！")

    # 4. 输入输出文件路径
    output_pdf = input("\n请输入新PDF保存路径（包含文件名，如：new.pdf）：").strip().replace('"', '')
    # 如果只输入目录，自动生成文件名
    if os.path.isdir(output_pdf):
        output_pdf = os.path.join(output_pdf, "替换后的PDF.pdf")

    # 5. 执行替换操作
    replace_pdf_pages(original_pdf, replacement_pdf, start_page, end_page, output_pdf)


if __name__ == "__main__":
    # 检查并安装PyPDF2库
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        print("⚠️  未找到PyPDF2库，正在自动安装...")
        import subprocess
        import sys

        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyPDF2", "-q"])
        print("✅ PyPDF2库安装完成！")
        from PyPDF2 import PdfReader, PdfWriter

    main()