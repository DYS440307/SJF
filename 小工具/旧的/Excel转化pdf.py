import subprocess
import os
import platform


def excel_to_pdf_libreoffice(excel_path, pdf_resolution=400):
    """
    无Excel依赖！使用LibreOffice将Excel转为PDF（支持所有工作表合并为一个PDF）

    参数:
        excel_path: Excel文件完整路径（.xlsx/.xls均支持）
        pdf_resolution: PDF分辨率（PPI），默认400
    """
    # 验证Excel文件是否存在
    if not os.path.exists(excel_path):
        print(f"❌ 错误：文件不存在 - {excel_path}")
        return

    # 获取输出PDF路径（与Excel同目录，同名）
    file_dir = os.path.dirname(excel_path)
    file_name = os.path.splitext(os.path.basename(excel_path))[0]
    pdf_path = os.path.join(file_dir, f"{file_name}.pdf")

    # 跳过已存在的PDF
    if os.path.exists(pdf_path):
        print(f"⚠️  已存在PDF文件，跳过转换：{pdf_path}")
        return

    # 1. 定位LibreOffice的soffice.exe路径（Windows默认路径）
    libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",  # 32位版本
        r"D:\Program Files\LibreOffice\program\soffice.exe"  # 自定义安装路径（可修改）
    ]
    soffice_path = None
    for path in libreoffice_paths:
        if os.path.exists(path):
            soffice_path = path
            break

    if not soffice_path:
        print("❌ 错误：未找到LibreOffice！请检查安装路径或手动指定soffice.exe路径")
        return

    # 2. 构建LibreOffice转换命令（无头模式）
    cmd = [
        soffice_path,
        "--headless",  # 无头模式（无GUI）
        "--norestore",  # 不恢复之前的文档，避免冲突
        "--invisible",  # 完全隐藏，不弹出窗口
        "--convert-to", f"pdf:calc_pdf_Export:{{\"PrintQuality\":{pdf_resolution}}}",  # 关键：设置400PPI分辨率
        "--outdir", file_dir,  # PDF输出目录
        excel_path  # 输入Excel文件
    ]

    try:
        print(f"🔄 正在转换（无Excel依赖）：{excel_path} -> {pdf_path}")
        print(f"📊 分辨率：{pdf_resolution}PPI")

        # 执行命令（隐藏命令行窗口，捕获输出）
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW  # Windows特有：隐藏命令行窗口
        )

        # 检查转换结果
        if result.returncode == 0 and os.path.exists(pdf_path):
            print(f"\n✅ 转换成功！PDF保存路径：{pdf_path}")
        else:
            print(f"\n❌ 转换失败！错误信息：")
            print(f"stdout: {result.stdout}")
            print(f"stderr: {result.stderr}")

    except Exception as e:
        print(f"\n❌ 转换异常：{str(e)}")


# 主程序执行
if __name__ == "__main__":
    # 目标Excel文件路径（原始字符串，避免转义）
    excel_file = r"E:\System\download\12301-500009焊锡段SOP-2024.8.19.xlsx"

    # 执行转换（400PPI分辨率）
    excel_to_pdf_libreoffice(excel_file, pdf_resolution=400)