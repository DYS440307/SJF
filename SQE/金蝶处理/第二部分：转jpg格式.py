import os
from wand.image import Image
from wand.color import Color

def process_single_pdf(pdf_path, dpi=500):
    """
    处理单个PDF：
    1. 提取所有页为高清白底图片
    2. 自动裁剪每页多余白边
    3. 拼接为一张长图输出
    4. 覆盖原PDF
    """
    try:
        file_dir, file_name = os.path.split(pdf_path)
        base_name = os.path.splitext(file_name)[0]
        img_path = os.path.join(file_dir, f"{base_name}.png")

        pages_images = []

        # 打开PDF，按页处理
        with Image(filename=pdf_path, resolution=dpi) as pdf:
            for page in pdf.sequence:
                with Image(page) as img:
                    # 转为白底
                    img.background_color = Color("white")
                    img.alpha_channel = 'remove'

                    # 自动裁剪
                    img.trim()

                    # 转成标准Image对象保存到列表
                    pages_images.append(img.clone())

        # 计算拼接后的总高度和最大宽度
        total_height = sum(img.height for img in pages_images)
        max_width = max(img.width for img in pages_images)

        # 创建一张新图用于拼接
        with Image(width=max_width, height=total_height, background=Color("white")) as final_img:
            y_offset = 0
            for img in pages_images:
                # 将每页贴上去
                final_img.composite(img, left=0, top=y_offset)
                y_offset += img.height

            # 保存最终拼接图
            final_img.save(filename=img_path)
            print(f"已生成长图：{img_path}")

        # 删除原PDF
        os.remove(pdf_path)
        print(f"已删除原PDF：{pdf_path}\n")

    except Exception as e:
        print(f"处理 {pdf_path} 时出错：{str(e)}\n")


def process_all_pdfs_in_folder(folder_path):
    """批量处理文件夹下所有PDF"""
    if not os.path.isdir(folder_path):
        print(f"错误：文件夹 {folder_path} 不存在")
        return

    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file_name)
            print(f"开始处理：{pdf_path}")
            process_single_pdf(pdf_path)

    print("所有PDF处理完毕")


if __name__ == "__main__":
    target_folder = r"E:\System\desktop\PY\BOMM"
    process_all_pdfs_in_folder(target_folder)
