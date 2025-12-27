import os
from PIL import Image
import pillow_heif

# 注册HEIC解码器，让Pillow支持HEIC格式
pillow_heif.register_heif_opener()


def heic_to_jpg(input_path, output_path, quality=80, scale_ratio=1.0):
    """
    将HEIC图片转换为JPG并压缩
    :param input_path: HEIC文件路径
    :param output_path: JPG输出路径
    :param quality: JPG压缩质量(1-95，越高质量越好，文件越大)
    :param scale_ratio: 尺寸缩放比例(0.1-1.0，1.0为原尺寸)
    """
    try:
        # 打开HEIC图片
        with Image.open(input_path) as img:
            # 缩放尺寸（可选）
            if scale_ratio != 1.0:
                width, height = img.size
                new_width = int(width * scale_ratio)
                new_height = int(height * scale_ratio)
                # 使用LANCZOS插值法缩放，画质更优（Pillow 9.1+推荐用Resampling.LANCZOS）
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # 转换为RGB模式（HEIC可能含透明通道，JPG不支持）
            if img.mode in ("RGBA", "P", "CMYK"):
                img = img.convert("RGB")

            # 保存为JPG并设置压缩质量
            img.save(output_path, "JPEG", quality=quality, optimize=True, progressive=True)
        print(f"✅ 转换成功：{output_path}")
    except Exception as e:
        print(f"❌ 转换失败 {input_path}：{str(e)}")


def batch_convert_heic(input_dir, output_dir, quality=80, scale_ratio=1.0):
    """
    批量转换文件夹中的HEIC文件
    :param input_dir: 输入文件夹路径
    :param output_dir: 输出文件夹路径
    :param quality: JPG压缩质量
    :param scale_ratio: 尺寸缩放比例
    """
    # 创建输出文件夹（如果不存在）
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 已创建输出文件夹：{output_dir}")

    # 遍历文件夹中的所有文件
    file_list = os.listdir(input_dir)
    if not file_list:
        print("⚠️ 输入文件夹中无文件")
        return

    heic_count = 0
    for filename in file_list:
        file_path = os.path.join(input_dir, filename)
        # 仅处理文件，跳过子文件夹
        if os.path.isfile(file_path):
            # 仅处理HEIC/HEIF格式文件（大小写兼容）
            if filename.lower().endswith((".heic", ".heif")):
                heic_count += 1
                # 构造输出文件名（替换后缀为jpg）
                jpg_filename = os.path.splitext(filename)[0] + ".jpg"
                output_path = os.path.join(output_dir, jpg_filename)
                # 转换文件
                heic_to_jpg(file_path, output_path, quality, scale_ratio)

    if heic_count == 0:
        print("⚠️ 输入文件夹中未找到HEIC/HEIF格式文件")
    else:
        print(f"\n🎉 批量转换完成，共处理 {heic_count} 个HEIC文件")


if __name__ == "__main__":
    # ************************* 配置参数 *************************
    # 你的原始HEIC文件所在路径（Windows路径用原始字符串r""避免转义）
    input_dir = r"Z:\3-品质部\实验室\邓洋枢\1-实验室相关文件\2-实验相关\2025年\十楼多媒体\UC000_周工更新过的质量需求书终版\过程资料\冷热冲击"
    # 输出路径：在原始路径下新建"转化后"文件夹
    output_dir = os.path.join(input_dir, "转化后")

    # 压缩配置（可根据需求调整）
    jpg_quality = 75  # JPG质量（1-95，建议70-85）
    scale_ratio = 0.8  # 尺寸缩放比例（1.0为原尺寸，0.8=80%）
    # ***********************************************************

    # 执行批量转换
    batch_convert_heic(input_dir, output_dir, jpg_quality, scale_ratio)