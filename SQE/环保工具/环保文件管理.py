import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse  # 兼容中文日期解析

# -------------------------- 全局配置项 --------------------------
# 替换为你的目标文件夹路径
TARGET_DIR = r'E:\System\download\厂商ROHS、REACH - 副本\4-一诺'
# 仅保留中文标注的匹配规则（完全忽略英文）
target_keys = {
    "客户名称": [
        # 仅匹配「报告抬头公司名称」中文标注，兼容冒号/空格/换行
        r"报告抬头公司名称\s*[:：]\s*([^\n]+)",
        r"报告抬头公司名称\s*\n\s*([^\n]+)"
    ],
    "样品名称": [
        # 仅匹配「样品名称」中文标注，兼容冒号/空格/换行
        r"样品名称\s*[:：]\s*([^\n]+)",
        r"样品名称\s*\n\s*([^\n]+)"
    ],
    "样品接收时间": [
        # 仅匹配「样品接收日期」中文标注，兼容冒号/空格/换行
        r"样品接收日期\s*[:：]\s*([^\n]+)",
        r"样品接收日期\s*\n\s*([^\n]+)"
    ]
}
# 报告有效期（天）
expire_days = 365
# 检测类型关键词（ROHS/REACH等，按需调整）
target_keywords = ["rohs", "reach", "pops", "svhc"]


# -------------------------- 工具函数 --------------------------
def filter_invalid_filename_chars(filename):
    """过滤文件名中的非法字符"""
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def clean_field_content(content):
    """清洗提取的中文字段内容"""
    if content == "未找到对应内容":
        return content
    # 去掉中英文冒号、多余空格，统一格式
    content = content.replace("：", "").replace(":", "") \
        .replace("，", ",").strip()
    # 合并多个连续空格为一个
    content = re.sub(r'\s+', ' ', content)
    return content


def calculate_expire_date(receive_date_str, days=365):
    """计算过期时间（仅解析中文日期，如2025.05.08/2025年05月08日）"""
    try:
        # 兼容中文日期格式解析
        receive_date = parse(receive_date_str, fuzzy=True)
        expire_date = receive_date + timedelta(days=days)
        # 统一输出为“XXXX年XX月XX日”格式
        return expire_date.strftime("%Y年%m月%d日")
    except Exception as e:
        print(f"⚠️ 日期解析失败：{receive_date_str}，错误：{e}")
        return "日期解析失败"


def get_unique_filename(file_dir, base_filename):
    """生成不重复的文件名（仅当文件重名时添加编号）"""
    filename_no_ext, ext = os.path.splitext(base_filename)
    unique_path = os.path.join(file_dir, base_filename)
    duplicate_num = 1
    # 仅文件存在时添加重名编号
    while os.path.exists(unique_path):
        new_filename = f"{filename_no_ext}_重名{duplicate_num}{ext}"
        unique_path = os.path.join(file_dir, new_filename)
        duplicate_num += 1
    return unique_path


# -------------------------- 核心提取函数（仅提取中文字段） --------------------------
def pdfplumber_extract_multi_page(pdf_path, target_keys, target_keywords):
    extract_result = {key: "未找到对应内容" for key in target_keys}
    extract_result["检测类型"] = ""
    matched_keywords = set()
    full_text = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历所有页面提取原生文本
            for page_num, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                if page_text:
                    full_text += f"\n【第{page_num}页】\n{page_text}"
                # 调试：打印第1页原始文本（方便排查提取问题）
                if page_num == 1:
                    print(f"\n【调试】{pdf_path} 第{page_num}页原始文本：\n{page_text}\n")

        # 无原生文本（扫描版）直接返回
        if not full_text.strip():
            print(f"⚠️ 该PDF无原生文本（可能是扫描版），无法提取字段")
            return extract_result

        # 仅匹配中文标注的字段
        for key, patterns in target_keys.items():
            if extract_result[key] == "未找到对应内容":
                for pattern in patterns:
                    # 仅匹配中文，关闭忽略大小写（中文无大小写）
                    match = re.search(pattern, full_text, re.MULTILINE | re.DOTALL)
                    if match:
                        extract_result[key] = match.group(1).strip()
                        break

        # 提取检测类型（兼容中英文关键词，但仅作为可选字段）
        full_text_lower = full_text.lower()
        for keyword in target_keywords:
            if keyword in full_text_lower:
                matched_keywords.add(keyword.upper())
        extract_result["检测类型"] = "/".join(matched_keywords) if matched_keywords else ""
        # 标记提取状态
        extract_result["找到内容的页码"] = "原生文本提取" if any(v != "未找到对应内容" for v in extract_result.values()) else "所有页均未找到"

    except Exception as e:
        extract_result = {"error": f"提取失败：{str(e)}"}

    return extract_result


# -------------------------- 单文件重命名函数 --------------------------
def rename_single_pdf(original_path):
    print(f"\n========== 开始处理文件：{original_path} ==========")

    # 1. 提取PDF内容（仅中文字段）
    extract_result = pdfplumber_extract_multi_page(original_path, target_keys, target_keywords)

    # 打印提取结果（清洗前）
    print("提取结果（清洗前）：")
    for key, value in extract_result.items():
        print(f"  {key}：{value}")

    # 2. 检查提取错误
    if "error" in extract_result:
        print(f"❌ 提取失败，跳过重命名：{extract_result['error']}")
        return False

    # 3. 清洗字段（仅保留中文内容）
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

    # 4. 检查必填中文字段（缺一不可）
    required_fields = [customer_name, sample_name, receive_date]
    if any(v == "未找到对应内容" for v in required_fields):
        print(f"❌ 关键必填中文字段缺失，跳过重命名")
        return False

    # 5. 计算过期时间
    expire_date = calculate_expire_date(receive_date, expire_days)
    if expire_date == "日期解析失败":
        print(f"❌ 过期时间计算失败，跳过重命名")
        return False

    # 6. 拼接文件名（仅中文核心字段）
    filename_parts = [
        customer_name,  # 报告抬头公司名称（中文）
        sample_name,    # 样品名称（中文）
        receive_date,   # 样品接收日期（中文）
        f"过期时间({expire_date})"  # 过期时间（中文格式）
    ]
    # 检测类型有值时追加（可选）
    if detect_type and detect_type.strip():
        filename_parts.append(detect_type)
    # 过滤空值，避免文件名混乱
    filename_parts = [part for part in filename_parts if part and part != "未找到对应内容"]
    base_filename = "_".join(filename_parts) + ".pdf"
    # 过滤非法字符
    base_filename = filter_invalid_filename_chars(base_filename)

    # 7. 生成唯一文件名
    original_dir = os.path.dirname(original_path)
    new_pdf_path = get_unique_filename(original_dir, base_filename)

    # 8. 执行重命名
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

    # 遍历目标文件夹下所有PDF
    for root, dirs, files in os.walk(target_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                total_count += 1
                file_path = os.path.join(root, file)
                # 处理单个PDF
                if rename_single_pdf(file_path):
                    success_count += 1
                else:
                    fail_count += 1
                    fail_files.append(file_path)

    # 打印批量处理汇总
    print("\n========== 批量处理完成 ==========")
    print(f"📊 汇总统计：")
    print(f"  总处理PDF数量：{total_count}")
    print(f"  ✅ 成功重命名：{success_count}")
    print(f"  ❌ 重命名失败：{fail_count}")

    # 打印失败文件列表
    if fail_files:
        print(f"\n❌ 失败的文件列表：")
        for fail_file in fail_files:
            print(f"  - {fail_file}")


# -------------------------- 主执行逻辑 --------------------------
if __name__ == "__main__":
    # 检查目标文件夹是否存在
    if not os.path.exists(TARGET_DIR):
        print(f"❌ 目标目录不存在：{TARGET_DIR}")
    else:
        # 启动批量处理
        batch_process_pdfs(TARGET_DIR)