import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse  # 兼容多语言/多格式日期解析

# -------------------------- 全局配置项 --------------------------
TARGET_DIR = r'E:\System\download\厂商ROHS、REACH - 副本\2-强升'
target_keys = {
    "客户名称": [
        r"报告抬头公司名称\s*([^\n]+)",  # 新模板核心（优先匹配）
        r"客户名称\s*([^\n]+)",  # 旧模板-中文
        r"Client Name\s*[:]?\s*([^\n]+)",  # 旧模板-英文（冒号可选）
        r"Company Name shown on Report\s*[:]?\s*([^\n]+)"  # 新模板英文
    ],
    "样品名称": [
        r"样品名称\s*([^\n]+)",  # 核心匹配（无冒号）
        r"Sample Name\s*[:]?\s*([^\n]+)"  # 英文（冒号可选）
    ],
    "样品接收时间": [
        r"样品接收日期\s*([^\n]+)",  # 新模板核心（无冒号）
        r"样品接收时间\s*([^\n]+)",  # 旧模板-中文
        r"Sample Received Date\s*[:]?\s*([^\n]+)",  # 新模板英文（冒号可选）
        r"Sample Receiving Date\s*[:]?\s*([^\n]+)"  # 旧模板英文
    ]
}
expire_days = 365
# 检测关键词：任意匹配、无顺序、遍历全页
target_keywords = ["rohs", "reach", "pops", "svhc"]


# -------------------------- 工具函数 --------------------------
def filter_invalid_filename_chars(filename):
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def clean_field_content(content):
    """清洗提取的字段内容：去掉中英文冒号、前后空白、多余空格，替换中文逗号为英文逗号"""
    if content == "未找到对应内容":
        return content
    # 去掉中英文冒号、多余空格，替换中文逗号为英文逗号（避免文件名乱码）
    content = content.replace("：", "").replace(":", "") \
        .replace("，", ",").strip()
    # 合并多个连续空格为一个
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


# -------------------------- 核心提取函数 --------------------------
def pdfplumber_extract_multi_page(pdf_path, target_keys, target_keywords):
    extract_result = {key: "未找到对应内容" for key in target_keys}
    extract_result["检测类型"] = ""
    # 收集所有匹配的检测关键词（去重）
    matched_keywords = set()

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 强制遍历PDF所有页面
            for page_num, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                if not page_text:
                    continue

                # 1. 提取基础信息（客户/样品/时间）：匹配到后不再重复提取
                for key, patterns in target_keys.items():
                    if extract_result[key] == "未找到对应内容":
                        for pattern in patterns:
                            match = re.search(pattern, page_text, re.IGNORECASE | re.MULTILINE)
                            if match:
                                extract_result[key] = match.group(1).strip()
                                break

                # 2. 提取检测类型：遍历全页+收集所有匹配的关键词（无顺序、去重）
                page_text_lower = page_text.lower()
                for keyword in target_keywords:
                    if keyword in page_text_lower:
                        matched_keywords.add(keyword.upper())  # 转大写并存入集合（自动去重）

        # 处理检测类型：将集合转为斜杠分隔的字符串（无顺序）
        if matched_keywords:
            extract_result["检测类型"] = "/".join(matched_keywords)
        else:
            extract_result["检测类型"] = ""

        # 记录找到基础信息的页码（仅用于日志）
        found_page = None
        for page_num, page in enumerate(pdf.pages, start=1):
            if all(v != "未找到对应内容" for v in
                   [extract_result["客户名称"], extract_result["样品名称"], extract_result["样品接收时间"]]):
                found_page = page_num
                break
        extract_result["找到内容的页码"] = found_page if found_page else "所有页均未找到"

    except Exception as e:
        extract_result = {"error": f"提取失败：{str(e)}"}

    return extract_result


# -------------------------- 单文件重命名函数 --------------------------
def rename_single_pdf(original_path):
    print(f"\n========== 开始处理文件：{original_path} ==========")

    # 1. 提取PDF内容
    extract_result = pdfplumber_extract_multi_page(original_path, target_keys, target_keywords)

    # 打印提取结果（清洗前）
    print("提取结果（清洗前）：")
    for key, value in extract_result.items():
        print(f"  {key}：{value}")

    # 2. 检查提取结果是否有错误
    if "error" in extract_result:
        print(f"❌ 提取失败，跳过重命名：{extract_result['error']}")
        return False

    # 3. 提取核心信息 + 清洗字段（删除客户参考信息相关）
    customer_name = clean_field_content(extract_result["客户名称"])
    sample_name = clean_field_content(extract_result["样品名称"])
    receive_date = clean_field_content(extract_result["样品接收时间"])
    detect_type = extract_result["检测类型"]

    # 打印清洗后的结果（删除客户参考信息行）
    print("提取结果（清洗后）：")
    print(f"  客户名称：{customer_name}")
    print(f"  样品名称：{sample_name}")
    print(f"  样品接收时间：{receive_date}")
    print(f"  检测类型：{detect_type}")

    # 4. 检查核心信息缺失（客户名称/样品名称/样品接收时间为必填）
    required_fields = [customer_name, sample_name, receive_date]
    if any(v == "未找到对应内容" for v in required_fields):
        print(f"❌ 关键必填信息缺失（客户名称/样品名称/样品接收时间），跳过重命名")
        return False

    # 5. 计算过期时间
    expire_date = calculate_expire_date(receive_date, expire_days)
    if expire_date == "日期解析失败":
        print(f"❌ 过期时间计算失败，跳过重命名")
        return False

    # 6. 拼接基础新文件名（删除客户参考信息拼接项）
    filename_parts = [
        customer_name,
        sample_name,
        receive_date,
        f"过期时间({expire_date})"
    ]
    # 检测类型有值才拼接
    if detect_type:
        filename_parts.append(detect_type)

    # 拼接所有部分，下划线分隔
    base_filename = "_".join(filename_parts) + ".pdf"
    # 过滤非法字符
    base_filename = filter_invalid_filename_chars(base_filename)

    # 7. 获取文件所在目录
    original_dir = os.path.dirname(original_path)

    # 8. 生成不重复文件名
    new_pdf_path = get_unique_filename(original_dir, base_filename)

    # 9. 执行重命名
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