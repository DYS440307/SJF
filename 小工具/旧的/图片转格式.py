import os
from PIL import Image
from pathlib import Path

# ===================== 配置参数（只需修改这部分） =====================
# 目标文件夹路径（Windows路径用原始字符串r""包裹，避免转义问题）
TARGET_FOLDER = r"Z:\3-品质部\实验室\邓洋枢\1-实验室相关文件\2-实验相关\2025年\十楼多媒体\UC000_周工更新过的质量需求书终版\过程资料\配件跌落\旧的"
# 转换后保存的子文件夹名（会自动创建，避免覆盖原文件）
OUTPUT_FOLDER = "JPG转换结果"
# 支持转换的源图片格式（小写，可根据需要添加）
SUPPORT_FORMATS = (".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".ico")
# JPG质量（0-100，85为通用最优值）
JPG_QUALITY = 85


# ===================== 核心转换逻辑 =====================
def convert_image_to_jpg(input_path, output_path):
    """
    单张图片转换为JPG格式
    :param input_path: 源图片路径
    :param output_path: 输出JPG路径
    """
    try:
        # 打开图片
        with Image.open(input_path) as img:
            # 处理透明背景（PNG/GIF等透明图转JPG时，透明区域填充白色）
            if img.mode in ("RGBA", "P"):
                # 创建白色背景画布
                bg = Image.new("RGB", img.size, (255, 255, 255))
                # 粘贴图片到背景上（保留Alpha通道）
                bg.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                img = bg
            # 转换为RGB模式（避免灰度图/索引图转换异常）
            if img.mode != "RGB":
                img = img.convert("RGB")
            # 保存为JPG
            img.save(output_path, "JPEG", quality=JPG_QUALITY, optimize=True)
        return True
    except Exception as e:
        print(f"❌ 转换失败 {input_path}：{str(e)}")
        return False


def batch_convert():
    """批量转换文件夹内的图片"""
    # 创建输出文件夹
    output_dir = Path(TARGET_FOLDER) / OUTPUT_FOLDER
    output_dir.mkdir(exist_ok=True)

    # 统计转换结果
    total = 0
    success = 0
    failed = 0

    # 遍历目标文件夹
    for file in Path(TARGET_FOLDER).iterdir():
        # 只处理文件 + 支持的格式
        if file.is_file() and file.suffix.lower() in SUPPORT_FORMATS:
            total += 1
            # 构建输出路径（保留原文件名，后缀改为jpg）
            output_file = output_dir / f"{file.stem}.jpg"
            # 执行转换
            if convert_image_to_jpg(str(file), str(output_file)):
                success += 1
                print(f"✅ 转换成功 {file.name} → {output_file.name}")
            else:
                failed += 1

    # 输出汇总信息
    print("\n" + "=" * 50)
    print(f"📊 转换完成 | 总计：{total} | 成功：{success} | 失败：{failed}")
    print(f"📁 转换后的文件保存在：{output_dir}")


if __name__ == "__main__":
    # 检查目标文件夹是否存在
    if not Path(TARGET_FOLDER).exists():
        print(f"❌ 错误：目标文件夹不存在 → {TARGET_FOLDER}")
    else:
        print(f"🚀 开始转换 {TARGET_FOLDER} 下的图片为JPG格式...")
        batch_convert()