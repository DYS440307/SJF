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

# ===================== 核心开关（手动修改这里即可） =====================
# True = 限制拼接后图片最大为20MB（会自动压缩）
# False = 不限制大小，直接最高质量保存（无任何压缩）
ENABLE_SIZE_LIMIT = True  # 核心开关，改这里！改这里！改这里！
TARGET_MAX_SIZE = 20       # 仅ENABLE_SIZE_LIMIT=True时生效，单位MB

# ===================== 控制台输入文件夹路径 =====================
if __name__ == "__main__":
    # 1. 控制台输入目标文件夹路径
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

    # ===================== 核心功能函数 =====================
    def get_valid_image_paths(folder):
        """遍历文件夹，获取所有有效图片路径（含HEIC格式）"""
        valid_formats = (".jpg", ".jpeg", ".png", ".bmp", ".heic", ".HEIC")
        valid_image_paths = []
        for fmt in valid_formats:
            valid_image_paths.extend(glob.glob(os.path.join(folder, f"*{fmt}")))
        # 去重+排序，保证顺序稳定
        return sorted(list(set(valid_image_paths)))

    def sanitize_filename(filename):
        """清理非法字符，避免保存失败"""
        illegal_chars = r'[\/:*?"<>|]'
        return re.sub(illegal_chars, '_', filename)[:50]

    def concat_images(image_paths):
        """拼接图片为长图（根据内置开关决定是否压缩）"""
        if not image_paths:
            print("错误：文件夹内未找到jpg/png/bmp/heic格式图片！")
            return None

        # 读取并缩放图片（统一宽度为2000px，避免过长）
        images = []
        base_width = 2000  # 固定基准宽度，简化逻辑
        for img_path in image_paths:
            try:
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

        # 创建拼接长图
        long_img = Image.new("RGB", (base_width, total_height), (255, 255, 255))
        current_y = 0
        for img in images:
            long_img.paste(img, (0, current_y))
            current_y += img.size[1]
            img.close()

        # 生成保存路径
        folder_name = sanitize_filename(os.path.basename(folder_path))
        save_path = os.path.join(folder_path, f"拼接长图_{folder_name}.jpg")

        # ===================== 核心逻辑：根据开关执行不同保存策略 =====================
        if ENABLE_SIZE_LIMIT:
            # 开启大小限制：压缩到指定MB内（默认20）
            print(f"🔒 已开启大小限制，目标最大{TARGET_MAX_SIZE}MB，开始压缩...")
            target_max_bytes = TARGET_MAX_SIZE * 1024 * 1024
            quality = 95
            while True:
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
                print(f"当前大小：{file_size / 1024 / 1024:.2f}MB > {TARGET_MAX_SIZE}MB，降低质量至{quality}")
        else:
            # 关闭大小限制：最高质量保存，无任何压缩
            print("🔓 已关闭大小限制，直接最高质量保存（无压缩）...")
            quality = 100  # 改为100（最高质量，原95）
            # 直接保存，无循环压缩、无大小检查
            long_img.save(
                save_path,
                format="JPEG",
                quality=quality,
                optimize=True,
                subsampling=0  # 关闭色度子采样，进一步提升质量
            )

        # 输出结果
        final_size = os.path.getsize(save_path) / 1024 / 1024
        print("=" * 50)
        print(f"✅ 拼接完成！")
        print(f"📁 保存路径：{save_path}")
        print(f"📏 文件大小：{final_size:.2f}MB（保存质量：{quality}）")
        long_img.close()
        return save_path

    # ===================== 执行核心逻辑 =====================
    # 获取有效图片
    image_paths = get_valid_image_paths(folder_path)
    print(f"✅ 找到 {len(image_paths)} 张有效图片（含HEIC格式）")

    # 拼接长图
    concat_images(image_paths)

    # 运行完成
    print("=" * 50)
    input("处理完成，按回车键退出...")