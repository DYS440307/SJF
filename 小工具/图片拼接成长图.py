from PIL import Image
import os
import sys
import glob
import io
import re
# 兼容新旧版本pillow-heif，注册HEIC格式解析器
try:
    from pillow_heif import HeifImagePlugin
    HeifImagePlugin.register()
except AttributeError:
    from pillow_heif import register_heif_opener
    register_heif_opener()

# ===================== 最顶部：控制台输入文件夹路径 =====================
if __name__ == "__main__":
    # 1. 控制台输入目标文件夹路径（核心入口，最顶部）
    print("=" * 50)
    folder_path = input("请输入要处理的图片文件夹路径：").strip()

    # 校验路径合法性
    if not os.path.exists(folder_path):
        print(f"错误：路径 {folder_path} 不存在！")
        input("按回车键退出...")
        sys.exit(1)
    if not os.path.isdir(folder_path):
        print(f"错误：{folder_path} 不是文件夹！")
        input("按回车键退出...")
        sys.exit(1)

    # ===================== 核心功能函数（修正所有检查问题） =====================
    def get_valid_image_paths(folder):
        """遍历文件夹，获取所有有效图片路径（含HEIC格式，解决变量隐藏问题）"""
        valid_formats = (".jpg", ".jpeg", ".png", ".bmp", ".heic", ".HEIC")
        # 重命名内部变量，避免与外部全局变量image_paths冲突（解决变量隐藏提示）
        valid_image_paths = []
        for fmt in valid_formats:
            valid_image_paths.extend(glob.glob(os.path.join(folder, f"*{fmt}")))
        # 去重+排序，保证顺序稳定
        return sorted(list(set(valid_image_paths)))

    def sanitize_filename(filename):
        """清理非法字符，避免保存失败"""
        illegal_chars = r'[\/:*?"<>|]'
        return re.sub(illegal_chars, '_', filename)[:50]

    def concat_images(image_paths, target_max_size=20):
        """拼接图片为长图，压缩到20MB内（兼容HEIC输入，解决long_img引用问题）"""
        if not image_paths:
            print("错误：文件夹内未找到jpg/png/bmp/heic格式图片！")
            return None

        # 读取并缩放图片（统一宽度为2000px，避免过长）
        images = []
        base_width = 2000  # 固定基准宽度，简化逻辑
        for img_path in image_paths:
            try:
                # 注册HEIC插件后，PIL可直接解析HEIC，无需额外修改读取逻辑
                img = Image.open(img_path).convert("RGB")
                w_percent = base_width / float(img.size[0])
                h_size = int(float(img.size[1]) * w_percent)
                h_size = min(h_size, 10000)  # 限制单张图高度
                img_resized = img.resize((base_width, h_size), Image.Resampling.LANCZOS)
                images.append(img_resized)
                img.close()
            except Exception as e:
                print(f"警告：跳过异常图片 {img_path}，错误：{str(e)[:50]}")
                continue

        if not images:
            print("错误：无有效图片可拼接！")
            return None

        # 检查总高度（避免超出PIL限制）
        total_height = sum(img.size[1] for img in images)
        max_height = 65500  # PIL对JPEG的高度上限
        if total_height > max_height:
            print(f"警告：拼接总高度({total_height})超上限，截断至{max_height}px")
            total_height = 0
            valid_images = []
            for img in images:
                if total_height + img.size[1] > max_height:
                    break
                valid_images.append(img)
                total_height += img.size[1]
            images = valid_images

        # 明确创建long_img变量，解决未解析引用提示
        long_img = Image.new("RGB", (base_width, total_height), (255, 255, 255))
        current_y = 0
        for img in images:
            long_img.paste(img, (0, current_y))
            current_y += img.size[1]
            img.close()

        # 生成保存路径（文件夹内命名为「拼接长图_文件夹名.jpg」）
        folder_name = sanitize_filename(os.path.basename(folder_path))
        save_path = os.path.join(folder_path, f"拼接长图_{folder_name}.jpg")

        # 压缩并保存（目标20MB）
        target_max_bytes = target_max_size * 1024 * 1024
        quality = 95
        while True:
            # 内存流缓存，避免文件句柄问题
            img_byte = io.BytesIO()
            long_img.save(img_byte, format="JPEG", quality=quality, optimize=True)
            img_byte.seek(0)

            # 写入文件
            with open(save_path, "wb") as f:
                f.write(img_byte.read())
            img_byte.close()

            # 检查大小
            file_size = os.path.getsize(save_path)
            if file_size <= target_max_bytes or quality <= 5:
                break
            quality -= 5
            print(f"当前大小：{file_size / 1024 / 1024:.2f}MB > 20MB，降低质量至{quality}")

        # 输出结果
        final_size = os.path.getsize(save_path) / 1024 / 1024
        print("=" * 50)
        print(f"✅ 拼接完成！")
        print(f"📁 保存路径：{save_path}")
        print(f"📏 文件大小：{final_size:.2f}MB（压缩质量：{quality}）")
        long_img.close()
        return save_path

    # ===================== 执行核心逻辑 =====================
    # 获取有效图片（外部变量名保持image_paths，内部已重命名无冲突）
    image_paths = get_valid_image_paths(folder_path)
    print(f"✅ 找到 {len(image_paths)} 张有效图片（含HEIC格式）")

    # 拼接长图
    concat_images(image_paths)

    # 运行完成
    print("=" * 50)
    input("处理完成，按回车键退出...")