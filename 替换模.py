from PyPDF2 import PdfReader, PdfWriter
import os
from datetime import datetime
import sys
import io

# 解决中文乱码问题（exe运行必备）
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def clean_file_path(file_path):
    """
    深度清理拖拽路径（兼容Windows所有拖拽场景）
    处理：引号、转义字符、多余空格、特殊符号
    """
    cleaned = file_path.strip()
    # 移除Windows拖拽自动添加的双引号/单引号
    if (cleaned.startswith('"') and cleaned.endswith('"')) or (cleaned.startswith("'") and cleaned.endswith("'")):
        cleaned = cleaned[1:-1]
    # 处理反斜杠转义（\\ → \）
    cleaned = cleaned.replace('\\\\', '\\')
    # 移除路径中间/前后的多余空格
    cleaned = ' '.join(cleaned.split())
    # 处理长路径或特殊字符路径
    if not os.path.exists(cleaned) and cleaned.startswith('C:'):
        # 尝试添加长路径前缀（Windows特殊处理）
        cleaned = f"\\\\?\\{cleaned}"
    return cleaned


def is_valid_pdf(file_path):
    """验证PDF文件有效性（存在+是PDF）"""
    if not file_path:
        return False, "文件路径不能为空"
    # 清理路径后验证
    file_path = clean_file_path(file_path)
    # 处理长路径前缀（如果添加了的话）
    if file_path.startswith("\\\\?\\") and os.path.exists(file_path[4:]):
        file_path = file_path[4:]
    if not os.path.exists(file_path):
        return False, f"文件不存在：\n{file_path}"
    if not os.path.isfile(file_path):
        return False, f"不是文件（可能是文件夹）：\n{file_path}"
    if not file_path.lower().endswith('.pdf'):
        return False, f"不是PDF文件（后缀错误）：\n{file_path}"
    return True, file_path


def get_pdf_pages(pdf_path):
    """获取PDF页数（带验证）"""
    try:
        valid, msg = is_valid_pdf(pdf_path)
        if not valid:
            raise ValueError(msg)
        reader = PdfReader(msg)  # msg是清理后的有效路径
        page_count = len(reader.pages)
        return True, page_count, msg
    except Exception as e:
        return False, str(e), None


def generate_output_path(original_path):
    """自动生成输出路径（原始文件名_年月日.pdf）"""
    original_path = clean_file_path(original_path)
    # 处理长路径前缀
    if original_path.startswith("\\\\?\\"):
        original_path = original_path[4:]
    dir_name = os.path.dirname(original_path)
    file_name = os.path.basename(original_path)
    name_no_ext = os.path.splitext(file_name)[0]
    today = datetime.now().strftime("%Y%m%d")
    output_name = f"{name_no_ext}_{today}.pdf"
    output_path = os.path.join(dir_name, output_name)
    # 重名处理（添加序号）
    counter = 1
    while os.path.exists(output_path):
        output_name = f"{name_no_ext}_{today}_{counter}.pdf"
        output_path = os.path.join(dir_name, output_name)
        counter += 1
    return output_path


def replace_pdf_pages(original_path, replace_path, start_page, end_page):
    """执行PDF页面替换"""
    try:
        # 验证原始PDF
        valid_orig, msg_orig = is_valid_pdf(original_path)
        if not valid_orig:
            raise ValueError(msg_orig)
        # 验证替换PDF
        valid_replace, msg_replace = is_valid_pdf(replace_path)
        if not valid_replace:
            raise ValueError(msg_replace)

        # 读取PDF
        orig_reader = PdfReader(msg_orig)
        replace_reader = PdfReader(msg_replace)
        orig_pages = len(orig_reader.pages)
        replace_pages = len(replace_reader.pages)

        # 验证页码范围
        if start_page < 1:
            raise ValueError(f"起始页不能小于1（当前输入：{start_page}）")
        if end_page < start_page:
            raise ValueError(f"结束页不能小于起始页（当前：{start_page} > {end_page}）")
        if end_page > orig_pages:
            raise ValueError(f"结束页超过原始PDF总页数（原始共{orig_pages}页，输入结束页：{end_page}）")

        # 验证替换PDF页数是否足够
        need_pages = end_page - start_page + 1
        if replace_pages < need_pages:
            raise ValueError(f"替换PDF页数不足！\n需要替换{need_pages}页，但替换PDF仅{replace_pages}页")

        # 写入新PDF
        writer = PdfWriter()
        # 1. 添加替换前的页面（1~start_page-1）
        for i in range(start_page - 1):
            writer.add_page(orig_reader.pages[i])
        # 2. 添加替换页面（取替换PDF的前need_pages页）
        for i in range(need_pages):
            writer.add_page(replace_reader.pages[i])
        # 3. 添加替换后的页面（end_page~末尾）
        for i in range(end_page, orig_pages):
            writer.add_page(orig_reader.pages[i])

        # 生成并保存输出文件
        output_path = generate_output_path(msg_orig)
        with open(output_path, "wb") as f:
            writer.write(f)

        return True, f"""
✅ 替换成功！

📋 操作详情：
• 原始PDF：{os.path.basename(msg_orig)}（{orig_pages}页）
• 替换PDF：{os.path.basename(msg_replace)}（{replace_pages}页）
• 替换范围：第{start_page}页 ~ 第{end_page}页（共{need_pages}页）
• 新文件路径：
{output_path}
"""
    except Exception as e:
        return False, f"❌ 替换失败：\n{str(e)}"


def main():
    # 界面提示（清晰告知支持拖拽）
    print("=" * 65)
    print("                  📄 PDF页面替换工具（exe版）")
    print("=" * 65)
    print("✅ 核心功能：识别PDF页数 + 替换指定页码 + 自动命名输出")
    print("✅ 操作说明：")
    print("   1. 可直接将PDF文件拖入输入框（自动识别路径）")
    print("   2. 页码从1开始计数（例如：替换第3-5页，起始页3，结束页5）")
    print("   3. 新文件自动保存在原始PDF同目录（文件名_年月日.pdf）")
    print("=" * 65)

    # 1. 输入原始PDF路径（支持拖拽）
    while True:
        print("\n📥 请输入【原始PDF】文件路径（可拖入文件）：")
        original_path = input("   → ").strip()
        if not original_path:
            print("❌ 路径不能为空，请重新输入！")
            continue
        # 验证原始PDF并获取页数
        success, result, valid_orig_path = get_pdf_pages(original_path)
        if success:
            print(f"✅ 识别成功：{os.path.basename(valid_orig_path)}（共{result}页）")
            original_path = valid_orig_path
            break
        else:
            print(f"❌ {result}，请重新输入！")

    # 2. 输入替换PDF路径（支持拖拽）
    while True:
        print("\n📥 请输入【替换PDF】文件路径（可拖入文件）：")
        replace_path = input("   → ").strip()
        if not replace_path:
            print("❌ 路径不能为空，请重新输入！")
            continue
        success, result, valid_replace_path = get_pdf_pages(replace_path)
        if success:
            print(f"✅ 识别成功：{os.path.basename(valid_replace_path)}（共{result}页）")
            replace_path = valid_replace_path
            break
        else:
            print(f"❌ {result}，请重新输入！")

    # 3. 输入替换页码（容错处理）
    while True:
        print("\n📝 请输入替换页码范围（从1开始）：")
        try:
            start_page = int(input("   起始页：").strip())
            end_page = int(input("   结束页（包含）：").strip())
            # 初步验证页码逻辑
            if start_page < 1:
                print("❌ 起始页不能小于1，请重新输入！")
                continue
            if end_page < start_page:
                print("❌ 结束页不能小于起始页，请重新输入！")
                continue
            # 这里不验证是否超过原始页数（留给replace函数统一处理）
            break
        except ValueError:
            print("❌ 请输入有效数字（不要输入文字/符号），请重新输入！")

    # 4. 确认并执行替换
    print(f"\n⚠️  即将执行替换：")
    print(f"   原始PDF：{os.path.basename(original_path)}")
    print(f"   替换范围：第{start_page}页 ~ 第{end_page}页")
    confirm = input("是否继续？（Y/n，默认Y）：").strip().lower()
    if confirm in ("n", "no"):
        print("\n❌ 操作已取消！")
        input("\n按Enter键退出...")
        return

    # 执行替换
    print("\n🔄 正在替换中，请稍候...")
    success, result = replace_pdf_pages(original_path, replace_path, start_page, end_page)
    print(result)
    input("\n按Enter键退出...")


if __name__ == "__main__":
    # 自动安装缺失的依赖（exe运行时如果缺失会报错，提前安装）
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        print("⚠️  检测到缺失依赖，正在自动安装PyPDF2...")
        import subprocess

        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "PyPDF2", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        print("✅ 依赖安装完成，正在重启程序...")
        # 重启程序以应用依赖
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()
    main()