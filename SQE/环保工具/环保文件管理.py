import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# -------------------------- 全局配置项 --------------------------
TARGET_DIR = r'E:\System\download\厂商ROHS、REACH - 副本\3-生湖\REACH'
# 配置Tesseract OCR路径（替换成你的安装路径）
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# 优化后的字段匹配规则（极致兼容英文模板排版）
target_keys = {
    "客户名称": [
        r"Company Name.*shown on Report[\s:]*\n?[\s:]*([^\n]+)",
        r"Company Name[\s\S]*?\n\s*([^\n]+)",
        r"客户名称\s*[:：]\s*([^\n]+)",
        r"报告抬头公司名称\s*([^\n]+)",
        r"Client Name\s*[:]?\s*([^\n]+)",
    ],
    "样品名称": [
        r"Sample Name[\s:]*\n?[\s:]*([^\n]+)",
        r"Sample Name[\s\S]*?\n\s*([^\n]+)",
        r"样品名称\s*[:：]\s*([^\n]+)",
    ],
    "样品接收时间": [
        r"Sample Received Date[\s:]*\n?[\s:]*([^\n]+)",
        r"Sample Received Date[\s\S]*?\n\s*([^\n]+)",
        r"收样日期\s*[:：]\s*([^\n]+)",
        r"样品接收日期\s*([^\n]+)",
        r"样品接收时间\s*([^\n]+)",
        r"Sample Receiving Date\s*[:]?\s*([^\n]+)",
    ]
}
expire_days = 365
target_keywords = ["rohs", "reach", "pops", "svhc"]
# OCR配置：识别语言（英文+中文）
OCR_LANG = 'eng+chi_sim'


# -------------------------- 工具函数 --------------------------
def filter_invalid_filename_chars(filename):
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def clean_field_content(content):
    if content == "未找到对应内容":
        return content
    content = content.replace("：", "").replace(":", "").replace("，", ",").strip()
    content = re.sub(r'\s+', ' ', content)
    return content


def calculate_expire_date(receive_date_str, days=365):
    try:
        receive_date = parse(receive_date_str, fuzzy=True)
        expire_date = receive_date + timedelta(days=days)
        return expire_date.strftime("%Y年%m月%d日")
    except Exception as e:
        print(f"⚠️ 日期解析失败：{receive_date_str}，错误：{e}")
        return "日期解析失败"


def get_unique_filename(file_dir, base_filename):
    filename_no_ext, ext = os.path.splitext(base_filename)
    unique_path = os.path.join(file_dir, base_filename)
    duplicate_num = 1
    while os.path.exists(unique_path):
        new_filename = f"{filename_no_ext}_重名{duplicate_num}{ext}"
        unique_path = os.path.join(file_dir, new_filename)
        duplicate_num += 1
    return unique_path


# -------------------------- 新增：OCR识别扫描版PDF文本 --------------------------
def ocr_scanned_pdf(pdf_path):
    """将扫描版PDF（图片）转成文本"""
    try:
        # 将PDF每页转成图片（分辨率300dpi保证识别精度）
        pages = convert_from_path(pdf_path, 300)
        full_text = ""
        for page_num, img in enumerate(pages, start=1):
            # 识别单页图片文本
            page_text = pytesseract.image_to_string(img, lang=OCR_LANG)
            full_text += f"\n【第{page_num}页】\n{page_text}"
            # 只识别前3页（多数报告关键信息在前3页），提升效率
            if page_num >= 3:
                break
        return full_text
    except Exception as e:
        print(f"⚠️ OCR识别失败：{e}")
        return ""


# -------------------------- 核心提取函数（兼容原生+扫描PDF） --------------------------
def pdf_extract_all(pdf_path, target_keys, target_keywords):
    extract_result = {key: "未找到对应内容" for key in target_keys}
    extract_result["检测类型"] = ""
    matched_keywords = set()
    full_text = ""

    # 第一步：尝试提取原生文本
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
    except:
        full_text = ""

    # 第二步：如果原生文本为空，说明是扫描版，用OCR识别
    if not full_text.strip():
        print(f"📌 检测到扫描版PDF，启动OCR识别...")
        full_text = ocr_scanned_pdf(pdf_path)

    # 调试打印识别到的文本
    print(f"\n【调试】最终识别到的文本：\n{full_text}\n")

    if not full_text:
        extract_result["error"] = "原生文本为空且OCR识别失败"
        return extract_result

    # 提取基础字段
    for key, patterns in target_keys.items():
        if extract_result[key] == "未找到对应内容":
            for pattern in patterns:
                match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
                if match:
                    extract_result[key] = match.group(1).strip()
                    break

    # 提取检测类型
    full_text_lower = full_text.lower()
    for keyword in target_keywords:
        if keyword in full_text_lower:
            matched_keywords.add(keyword.upper())
    extract_result["检测类型"] = "/".join(matched_keywords) if matched_keywords else ""
    extract_result["找到内容的页码"] = "OCR识别/原生文本提取"

    return extract_result


# -------------------------- 单文件重命名函数 --------------------------
def rename_single_pdf(original_path):
    print(f"\n========== 开始处理文件：{original_path} ==========")

    # 1. 提取PDF内容（兼容原生+扫描）
    extract_result = pdf_extract_all(original_path, target_keys, target_keywords)

    # 打印提取结果（清洗前）
    print("提取结果（清洗前）：")
    for key, value in extract_result.items():
        print(f"  {key}：{value}")

    # 2. 检查提取结果是否有错误
    if "error" in extract_result:
        print(f"❌ 提取失败，跳过重命名：{extract_result['error']}")
        return False

    # 3. 提取核心信息 + 清洗字段
    customer_name = clean_field_content(extract_result["客户名称"])
    sample_name = clean_field_content(extract_result["样品名称"])
    receive_date = clean_field_content(extract_result["样品接收时间"])
    detect_type = extract_result["检测类型"]

    # 打印清洗后的结果
    print("提取结果（清洗后）：")
    print(f"  客户名称：{customer_name}")
    print(f"  样品名称：{sample_name}")
    print(f"  样品接收时间：{receive_date}")
    print(f"  检测类型：{detect_type}")

    # 4. 检查核心信息缺失
    required_fields = [customer_name, sample_name, receive_date]
    if any(v == "未找到对应内容" for v in required_fields):
        print(f"❌ 关键必填信息缺失，跳过重命名")
        return False

    # 5. 计算过期时间
    expire_date = calculate_expire_date(receive_date, expire_days)
    if expire_date == "日期解析失败":
        print(f"❌ 过期时间计算失败，跳过重命名")
        return False

    # 6. 拼接文件名
    filename_parts = [customer_name, sample_name, receive_date, f"过期时间({expire_date})"]
    if detect_type:
        filename_parts.append(detect_type)
    base_filename = "_".join(filename_parts) + ".pdf"
    base_filename = filter_invalid_filename_chars(base_filename)

    # 7. 生成不重复文件名并执行重命名
    original_dir = os.path.dirname(original_path)
    new_pdf_path = get_unique_filename(original_dir, base_filename)
    try:
        os.rename(original_path, new_pdf_path)
        print(f"✅ 重命名成功！新路径：{new_pdf_path}")
        return True
    except Exception as e:
        print(f"❌ 重命名失败：{str(e)}")
        return False


# -------------------------- 批量处理函数 --------------------------
def batch_process_pdfs(target_dir):
    total_count = 0
    success_count = 0
    fail_count = 0
    fail_files = []

    for root, dirs, files in os.walk(target_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                total_count += 1
                file_path = os.path.join(root, file)
                if rename_single_pdf(file_path):
                    success_count += 1
                else:
                    fail_count += 1
                    fail_files.append(file_path)

    print("\n========== 批量处理完成 ==========")
    print(f"📊 汇总统计：")
    print(f"  总处理PDF数量：{total_count}")
    print(f"  ✅ 成功重命名：{success_count}")
    print(f"  ❌ 重命名失败：{fail_count}")

    if fail_files:
        print(f"\n❌ 失败的文件列表：")
        for fail_file in fail_files:
            print(f"  - {fail_file}")


# -------------------------- 主执行逻辑 --------------------------
if __name__ == "__main__":
    if not os.path.exists(TARGET_DIR):
        print(f"❌ 目标目录不存在：{TARGET_DIR}")
    else:
        batch_process_pdfs(TARGET_DIR)