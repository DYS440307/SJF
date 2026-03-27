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

# ===================== 核心配置（手动修改这里即可，注释清晰！） =====================
# 1. 大小限制开关（原功能）
ENABLE_SIZE_LIMIT = True  # True=压缩到指定MB，False=最高质量无压缩
TARGET_MAX_SIZE = 20       # 仅ENABLE_SIZE_LIMIT=True生效，单位MB
# 2. A4横版拼接专属配置（重点！改这些调整比例/留白/大小）
BASE_DPI = 300             # A4的DPI，72(屏幕)/150(低精度打印)/300(高精度打印)
A4_COLOR = (255, 255, 255) # A4画布背景色，RGB格式（如(240,240,240)是浅灰）
BASE_ROW_HEIGHT = 600      # 每行图片的基础高度（核心！调大=图片整体变大，调小=能放更多张）
IMG_SPACING = 10           # 图片之间的间距（像素，调小=更少空白，0=无缝拼接）
CANVAS_MARGIN = 15         # 画布整体的内边距（像素，调小=画布利用更充分）

# ===================== 工具函数（无需修改，复用逻辑） =====================
def get_valid_image_paths(folder):
    """遍历文件夹，获取所有有效图片路径（含HEIC格式）"""
    valid_formats = (".jpg", ".jpeg", ".png", ".bmp", ".heic", ".HEIC")
    valid_image_paths = []
    for fmt in valid_formats:
        valid_image_paths.extend(glob.glob(os.path.join(folder, f"*{fmt}")))
    return sorted(list(set(valid_image_paths)))  # 去重+排序

def sanitize_filename(filename):
    """清理非法字符，避免保存失败"""
    illegal_chars = r'[\/:*?"<>|]'
    return re.sub(illegal_chars, '_', filename)[:50]

def save_image(img, save_path):
    """复用保存逻辑（统一大小限制/最高质量，减少冗余）"""
    if ENABLE_SIZE_LIMIT:
        print(f"🔒 已开启大小限制，目标最大{TARGET_MAX_SIZE}MB，开始压缩...")
        target_max_bytes = TARGET_MAX_SIZE * 1024 * 1024
        quality = 95
        while True:
            img_byte = io.BytesIO()
            img.save(img_byte, format="JPEG", quality=quality, optimize=True)
            img_byte.seek(0)
            with open(save_path, "wb") as f:
                f.write(img_byte.read())
            img_byte.close()
            file_size = os.path.getsize(save_path)
            if file_size <= target_max_bytes or quality <= 5:
                break
            quality -= 5
            print(f"当前大小：{file_size / 1024 / 1024:.2f}MB > {TARGET_MAX_SIZE}MB，降低质量至{quality}")
    else:
        print("🔓 已关闭大小限制，直接最高质量保存（无压缩）...")
        quality = 100
        img.save(
            save_path,
            format="JPEG",
            quality=quality,
            optimize=True,
            subsampling=0  # 关闭色度子采样，提升质量
        )
    final_size = os.path.getsize(save_path) / 1024 / 1024
    return final_size, quality

# ===================== 原有功能：竖版长图拼接（无修改） =====================
def concat_long_image(image_paths, folder_name, save_root):
    """原有逻辑：拼接为竖版长图"""
    if not image_paths:
        print("错误：文件夹内未找到有效图片！")
        return None
    images = []
    base_width = 2000  # 固定基准宽度
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
    # 检查PIL高度上限
    total_height = sum(img.size[1] for img in images)
    max_height = 65500
    if total_height > max_height:
        print(f"警告：拼接总高度({total_height})超上限，截断至{max_height}px")
        total_height, valid_images = 0, []
        for img in images:
            if total_height + img.size[1] > max_height:
                break
            valid_images.append(img)
            total_height += img.size[1]
        images = valid_images
    # 创建长图
    long_img = Image.new("RGB", (base_width, total_height), A4_COLOR)
    current_y = 0
    for img in images:
        long_img.paste(img, (0, current_y))
        current_y += img.size[1]
        img.close()
    # 保存
    save_path = os.path.join(save_root, f"拼接长图_{folder_name}.jpg")
    final_size, quality = save_image(long_img, save_path)
    print("=" * 60)
    print(f"✅ 长图拼接完成！")
    print(f"📁 保存路径：{save_path}")
    print(f"📏 文件大小：{final_size:.2f}MB（保存质量：{quality}）")
    long_img.close()
    return save_path

# ===================== 优化后核心：A4横版紧凑拼接（无多余空白+可调整比例） =====================
def concat_a4_horizontal(image_paths, folder_name, save_root):
    """A4横版拼接：按行填充紧凑排版，图片按自身比例适配，可调整间距/大小，大幅减少空白"""
    if not image_paths:
        print("错误：文件夹内未找到有效图片！")
        return None
    # 1. 定义A4横版标准像素尺寸（关键！宽高互换：297mm(宽)×210mm(高)，按DPI换算）
    a4_mm_w, a4_mm_h = 297, 210  # 横向A4：宽297mm，高210mm
    a4_px_w = int(a4_mm_w * BASE_DPI / 25.4)  # mm转像素：×DPI÷25.4
    a4_px_h = int(a4_mm_h * BASE_DPI / 25.4)
    print(f"📄 A4横版画布（{BASE_DPI}DPI）：{a4_px_w}×{a4_px_h} 像素")
    # 计算画布实际可用区域（扣除整体边距）
    usable_w = a4_px_w - 2 * CANVAS_MARGIN  # 横向可用宽度
    usable_h = a4_px_h - 2 * CANVAS_MARGIN  # 纵向可用高度

    # 2. 读取并预处理图片（转RGB，记录宽高比，异常图片直接跳过）
    img_info_list = []  # 存储(图片对象, 宽高比)
    for img_path in image_paths:
        try:
            img = Image.open(img_path).convert("RGB")
            w, h = img.size
            ratio = w / h  # 图片原始宽高比（核心，用于按比例分配宽度）
            img_info_list.append((img, ratio))
        except Exception as e:
            print(f"警告：跳过异常图片 {img_path}，错误：{str(e)[:50]}")
            continue
    if not img_info_list:
        print("错误：无有效图片可拼接！")
        return None
    img_count = len(img_info_list)
    print(f"📸 参与A4横版拼接的图片数量：{img_count} 张")

    # 3. 核心：按行填充紧凑排版逻辑
    rows = []  # 存储每行的图片信息：[(img, ratio), ...]
    current_row = []  # 当前行的图片信息
    current_total_ratio = 0  # 当前行的总宽高比
    for img, ratio in img_info_list:
        # 临时加入当前行，计算总比例
        temp_total_ratio = current_total_ratio + ratio
        # 若当前行加入后仍能放下，直接加入；否则换行
        current_row.append((img, ratio))
        current_total_ratio = temp_total_ratio
        # 预判：若当前行总宽度超过可用宽度，最后一张移到下一行
        # 每行宽度=总比例×行高 + 图片间距×(图片数-1)
        predict_width = current_total_ratio * BASE_ROW_HEIGHT + IMG_SPACING * (len(current_row)-1)
        if predict_width > usable_w:
            # 移除最后一张，当前行定型，开始新行
            last_img, last_ratio = current_row.pop()
            current_total_ratio -= last_ratio
            if current_row:  # 避免空行
                rows.append(current_row)
            # 新行初始化
            current_row = [(last_img, last_ratio)]
            current_total_ratio = last_ratio
    # 把最后一行加入
    if current_row:
        rows.append(current_row)
    print(f"📐 自动紧凑排版：共{len(rows)}行（无多余空白）")

    # 4. 创建A4横版画布，逐行粘贴图片
    a4_img = Image.new("RGB", (a4_px_w, a4_px_h), A4_COLOR)
    current_y = CANVAS_MARGIN  # 纵向起始坐标（扣除上边距）
    for row in rows:
        row_img_count = len(row)
        row_total_ratio = sum(ratio for _, ratio in row)
        # 计算每行实际行高（适配纵向可用区域，防止超出画布）
        row_actual_h = min(BASE_ROW_HEIGHT, (usable_h - (len(rows)-1)*IMG_SPACING) // len(rows))
        # 计算每张图片的实际宽度（按总比例分配，保证宽高比）
        # 可用宽度=画布可用宽 - 图片间距×(图片数-1)
        row_usable_w = usable_w - IMG_SPACING * (row_img_count - 1)
        each_img_base_w = row_usable_w / row_total_ratio
        # 逐张粘贴当前行的图片
        current_x = CANVAS_MARGIN  # 横向起始坐标（扣除左边距）
        for img, ratio in row:
            # 按比例计算图片实际宽高（无拉伸，不变形）
            img_actual_w = int(each_img_base_w * ratio)
            img_actual_h = row_actual_h
            # 高质量缩放图片
            img_resized = img.resize((img_actual_w, img_actual_h), Image.Resampling.LANCZOS)
            # 粘贴图片（左对齐，紧凑排列）
            a4_img.paste(img_resized, (current_x, current_y))
            # 更新横向坐标（图片宽度+间距）
            current_x += img_actual_w + IMG_SPACING
            # 关闭临时图片，释放内存
            img.close()
            img_resized.close()
        # 更新纵向坐标（行高+间距）
        current_y += row_actual_h + IMG_SPACING

    # 5. 生成保存路径（区分长图，标注横版紧凑）
    save_path = os.path.join(save_root, f"A4横版紧凑拼接_{folder_name}.jpg")
    # 6. 复用保存逻辑（大小限制/最高质量）
    final_size, quality = save_image(a4_img, save_path)

    # 输出结果
    print("=" * 60)
    print(f"✅ A4横版紧凑拼接完成！")
    print(f"📁 保存路径：{save_path}")
    print(f"📏 文件大小：{final_size:.2f}MB（保存质量：{quality}）")
    a4_img.close()
    return save_path

# ===================== 主程序入口（模式选择：长图/A4横版） =====================
if __name__ == "__main__":
    # 控制台欢迎信息
    print("=" * 60)
    print("📷 图片拼接工具 V2.0 | 长图拼接/A4横版紧凑拼接 | 兼容HEIC")
    print("💡 可修改顶部「核心配置」调整A4横版的图片大小/间距/留白")
    print("=" * 60)
    # 输入并校验文件夹路径
    folder_path = input("请输入要处理的图片文件夹路径：").strip()
    if not os.path.exists(folder_path):
        print(f"错误：路径 {folder_path} 不存在！")
        input("按回车键退出...")
        sys.exit(1)
    if not os.path.isdir(folder_path):
        print(f"错误：{folder_path} 不是文件夹！")
        input("按回车键退出...")
        sys.exit(1)
    # 扫描有效图片
    image_paths = get_valid_image_paths(folder_path)
    if not image_paths:
        print("错误：文件夹内未找到jpg/png/bmp/heic格式图片！")
        input("按回车键退出...")
        sys.exit(1)
    print(f"✅ 扫描完成，找到 {len(image_paths)} 张有效图片（含HEIC格式）")
    # 选择拼接模式
    print("\n📌 请选择拼接模式：")
    print("  1 - 传统竖版长图拼接（原功能，无限滚动）")
    print("  2 - A4横版紧凑拼接（无多余空白，可调整图片比例/间距）")
    while True:
        choice = input("请输入数字1或2选择模式：").strip()
        if choice in ["1", "2"]:
            break
        print("❌ 输入错误！请仅输入数字1或2")
    # 初始化参数，执行对应逻辑
    folder_name = sanitize_filename(os.path.basename(folder_path))
    if choice == "1":
        concat_long_image(image_paths, folder_name, folder_path)
    else:
        concat_a4_horizontal(image_paths, folder_name, folder_path)
    # 运行完成
    print("=" * 60)
    input("处理完成，按回车键退出...")