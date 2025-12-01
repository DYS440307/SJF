from PyPDF2 import PdfReader, PdfWriter
import os
from datetime import datetime
import sys
import io

# 1. 解决中文乱码+控制台拖拽支持
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def clean_drag_path(file_path):
    """
    终极拖拽路径清理（覆盖所有Windows拖拽场景）
    处理：引号、转义符、长路径、网络路径、多余空格
    """
    if not file_path:
        return ""

    cleaned = file_path.strip()

    # 关键：移除Windows拖拽强制添加的双引号（最常见问题）
    if cleaned.startswith('"') and cleaned.endswith('"'):
        cleaned = cleaned[1:-1]
    if cleaned.startswith("'") and cleaned.endswith("'"):
        cleaned = cleaned[1:-1]

    # 处理反斜杠转义（\\ → \）
    cleaned = cleaned.replace('\\\\', '\\')

    # 移除路径中间/前后的多余空格（拖拽时可能误选空格）
    cleaned = ' '.join(cleaned.split())

    # 处理Windows长路径（超过260字符）
    if len(cleaned) > 256 and not cleaned.startswith("\\\\?\\"):
        cleaned = f"\\\\?\\{cleaned}"

    # 处理网络路径（// → \\）
    if cleaned.startswith("//"):
        cleaned = cleaned.replace("//", "\\\\", 1)

    # 验证路径存在性（如果不存在，尝试去掉长路径前缀再验证）
    if not os.path.exists(cleaned) and cleaned.startswith("\\\\?\\"):
        cleaned = cleaned[4:]

    return cleaned


def is_valid_pdf(file_path):
    """验证PDF有效性，拖拽后自动校验"""
    if not file_path:
        return False, "路径不能为空"

    cleaned_path = clean_drag_path(file_path)

    # 检查文件是否存在
    if not os.path.exists(cleaned_path):
        return False, f"文件不存在：\n{cleaned_path}"

    # 检查是否是文件（不是文件夹）
    if not os.path.isfile(cleaned_path):
        return False, f"这是文件夹，不是文件：\n{cleaned_path}"

    # 检查是否是PDF
    if not cleaned_path.lower().endswith('.pdf'):
        return False, f"不是PDF文件（支持后缀：.pdf）：\n{cleaned_path}"

    return True, cleaned_path


def get_pdf_page_count(pdf_path):
    """获取PDF页数，带拖拽验证"""
    try:
        valid, msg = is_valid_pdf(pdf_path)
        if not valid:
            raise ValueError(msg)
        reader = PdfReader(msg)
        return len(reader.pages), msg  # 返回页数和清理后的有效路径
    except Exception as e:
        return None, str(e)


def generate_output_path(original_path):
    """自动生成输出路径（原始文件同目录+日期）"""
    original_path = clean_drag_path(original_path)
    if original_path.startswith("\\\\?\\"):
        original_path = original_path[4:]
    dir_name = os.path.dirname(original_path)
    file_name = os.path.basename(original_path)
    name_no_ext = os.path.splitext(file_name)[0]
    today = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(dir_name, f"{name_no_ext}_{today}.pdf")

    # 重名处理
    counter = 1
    while os.path.exists(output_path):
        output_path = os.path.join(dir_name, f"{name_no_ext}_{today}_{counter}.pdf")
        counter += 1
    return output_path


def replace_pdf(original_path, replace_path, start_page, end_page):
    """执行替换逻辑"""
    try:
        # 验证原始PDF
        orig_valid, orig_path = is_valid_pdf(original_path)
        if not orig_valid:
            return False, orig_path
        # 验证替换PDF
        repl_valid, repl_path = is_valid_pdf(replace_path)
        if not repl_valid:
            return False, repl_path

        # 读取PDF
        orig_reader = PdfReader(orig_path)
        repl_reader = PdfReader(repl_path)
        orig_page_num = len(orig_reader.pages)
        repl_page_num = len(repl_reader.pages)

        # 验证页码范围
        if start_page < 1:
            return False, f"起始页不能小于1（输入：{start_page}）"
        if end_page < start_page:
            return False, f"结束页不能小于起始页（输入：{start_page} > {end_page}）"
        if end_page > orig_page_num:
            return False, f"结束页超过原始PDF总页数（原始共{orig_page_num}页，输入：{end_page}）"

        # 验证替换PDF页数是否足够
        need_pages = end_page - start_page + 1
        if repl_page_num < need_pages:
            return False, f"替换PDF页数不足（需要{need_pages}页，仅{repl_page_num}页）"

        # 写入新PDF
        writer = PdfWriter()
        # 添加替换前页面
        for i in range(start_page - 1):
            writer.add_page(orig_reader.pages[i])
        # 添加替换页面
        for i in range(need_pages):
            writer.add_page(repl_reader.pages[i])
        # 添加替换后页面
        for i in range(end_page, orig_page_num):
            writer.add_page(orig_reader.pages[i])

        # 保存文件
        output_path = generate_output_path(orig_path)
        with open(output_path, "wb") as f:
            writer.write(f)

        return True, f"""
✅ 替换成功！

📋 操作详情：
• 原始PDF：{os.path.basename(orig_path)}（{orig_page_num}页）
• 替换PDF：{os.path.basename(repl_path)}（{repl_page_num}页）
• 替换范围：第{start_page}页 ~ 第{end_page}页（共{need_pages}页）
• 新文件路径：
{output_path}
"""
    except Exception as e:
        return False, f"❌ 替换失败：\n{str(e)}"


def main():
    print("=" * 70)
    print("                  📄 PDF页面替换工具（拖拽专用版）")
    print("=" * 70)
    print("✅ 拖拽说明：直接将PDF文件拖入黑框，松开鼠标后按回车即可！")
    print("✅ 页码说明：从1开始计数（例：替换第3-5页 → 起始3，结束5）")
    print("✅ 输出说明：新文件自动保存在原始PDF同目录（文件名_年月日.pdf）")
    print("=" * 70)

    # 1. 拖拽/输入原始PDF
    while True:
        print("\n📥 请将【原始PDF】拖入黑框，或手动输入路径后按回车：")
        original_path = input("   → ").strip()
        page_count, msg = get_pdf_page_count(original_path)
        if page_count:
            print(f"✅ 识别成功：{os.path.basename(msg)}（共{page_count}页）")
            original_path = msg
            break
        else:
            print(f"❌ 错误：{msg}，请重新操作！")

    # 2. 拖拽/输入替换PDF
    while True:
        print("\n📥 请将【替换PDF】拖入黑框，或手动输入路径后按回车：")
        replace_path = input("   → ").strip()
        page_count, msg = get_pdf_page_count(replace_path)
        if page_count:
            print(f"✅ 识别成功：{os.path.basename(msg)}（共{page_count}页）")
            replace_path = msg
            break
        else:
            print(f"❌ 错误：{msg}，请重新操作！")

    # 3. 输入页码范围
    while True:
        print("\n📝 请输入替换页码范围：")
        try:
            start_page = int(input("   起始页：").strip())
            end_page = int(input("   结束页（包含）：").strip())
            if start_page < 1:
                print("❌ 起始页不能小于1，请重新输入！")
                continue
            if end_page < start_page:
                print("❌ 结束页不能小于起始页，请重新输入！")
                continue
            break
        except ValueError:
            print("❌ 请输入有效数字（不要输入文字/符号），请重新输入！")

    # 4. 确认并执行
    print(f"\n⚠️  即将替换：{os.path.basename(original_path)} 的第{start_page}-{end_page}页")
    confirm = input("是否继续？（Y/n，默认Y）：").strip().lower()
    if confirm in ("n", "no"):
        print("\n❌ 操作已取消！")
        input("\n按Enter键退出...")
        return

    print("\n🔄 正在替换中，请稍候...")
    success, result = replace_pdf(original_path, replace_path, start_page, end_page)
    print(result)
    input("\n按Enter键退出...")


if __name__ == "__main__":
    # 自动安装依赖
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        print("⚠️  缺失依赖，正在自动安装PyPDF2...")
        import subprocess

        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "PyPDF2", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        print("✅ 依赖安装完成，正在重启...")
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()
    main()