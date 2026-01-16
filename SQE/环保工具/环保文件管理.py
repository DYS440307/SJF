import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse  # 兼容多语言/多格式日期解析

# -------------------------- 全局配置项 --------------------------
# 目标处理目录（所有PDF都在这个目录下，含子目录）
TARGET_DIR = r'E:\System\download\厂商ROHS、REACH - 副本\1-诚意达\REACH'
# 目标提取项：支持中英文两套关键词（正则匹配大小写不敏感）
target_keys = {
    "客户名称": [
        r"客户名称[:：]\s*([^\n]+)",  # 中文关键词正则
        r"Client Name[:]\s*([^\n]+)"  # 英文关键词正则（Client Name: 后内容）
    ],
    "样品名称": [
        r"样品名称[:：]\s*([^\n]+)",  # 中文关键词正则
        r"Sample Name[:]\s*([^\n]+)"  # 英文关键词正则（Sample Name: 后内容）
    ],
    "样品接收时间": [
        r"样品接收时间[:：]\s*([^\n]+)",  # 中文关键词正则
        r"Sample Receiving Date[:]\s*([^\n]+)"  # 英文关键词正则（Sample Receiving Date: 后内容）
    ]
}
# 过期时间偏移量（365天）
expire_days = 365
# 要查找的关键字（大小写不敏感）
target_keywords = ["rohs", "reach", "svhc"]  # 新增svhc适配英文报告


# -------------------------- 工具函数 --------------------------
def filter_invalid_filename_chars(filename):
    """过滤文件名中的非法字符（Windows系统）"""
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def calculate_expire_date(receive_date_str, days=365):
    """
    兼容中英文日期格式的过期时间计算
    支持：2025年05月13日、Jun 21, 2024、2025.5.13等格式
    """
    try:
        # 用dateutil自动识别日期格式（兼容中英文）
        receive_date = parse(receive_date_str, fuzzy=True)
        # 计算过期时间
        expire_date = receive_date + timedelta(days=days)

        # 匹配原日期格式，保持输出格式一致
        # 中文日期（2025年05月13日）
        if re.match(r"\d{4}年\d{1,2}月\d{1,2}日", receive_date_str):
            return expire_date.strftime("%Y年%m月%d日")
        # 英文日期（Jun 21, 2024）
        elif re.match(r"[A-Za-z]{3} \d{1,2}, \d{4}", receive_date_str):
            return expire_date.strftime("%b %d, %Y")
        # 其他格式（如2025.5.13）
        else:
            return expire_date.strftime("%Y-%m-%d")
    except Exception as e:
        print(f"⚠️ 日期解析失败：{receive_date_str}，错误：{e}")
        return "日期解析失败"


# -------------------------- 核心提取函数 --------------------------
def pdfplumber_extract_multi_page(pdf_path, target_keys, target_keywords):
    """
    多页遍历提取PDF内容（兼容中英文模板）
    优先匹配中文关键词，匹配不到则匹配英文关键词
    """
    extract_result = {key: "未找到对应内容" for key in target_keys}
    extract_result["检测类型"] = ""  # 存储找到的RoHs/REACH/SVHC关键字
    found_page = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历所有页面（页码从1开始）
            for page_num, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                if not page_text:
                    continue  # 该页无文本，跳过

                # 1. 提取核心信息（中英文关键词都尝试）
                for key, patterns in target_keys.items():
                    if extract_result[key] == "未找到对应内容":
                        for pattern in patterns:
                            match = re.search(pattern, page_text, re.IGNORECASE)  # 大小写不敏感
                            if match:
                                extract_result[key] = match.group(1).strip()
                                break  # 匹配到一个就停止

                # 2. 查找检测类型关键字（ROHS/REACH/SVHC，大小写不敏感）
                if not extract_result["检测类型"]:
                    page_text_lower = page_text.lower()
                    for keyword in target_keywords:
                        if keyword in page_text_lower:
                            extract_result["检测类型"] = keyword.upper()
                            break

                # 基础信息全找到就终止遍历
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
    """处理单个PDF文件的重命名（兼容中英文模板），返回处理结果"""
    print(f"\n========== 开始处理文件：{original_path} ==========")

    # 1. 提取PDF内容
    extract_result = pdfplumber_extract_multi_page(original_path, target_keys, target_keywords)

    # 打印提取结果
    print("提取结果：")
    for key, value in extract_result.items():
        print(f"  {key}：{value}")

    # 2. 检查提取结果是否有错误
    if "error" in extract_result:
        print(f"❌ 提取失败，跳过重命名：{extract_result['error']}")
        return False

    # 3. 提取核心信息
    customer_name = extract_result["客户名称"]
    sample_name = extract_result["样品名称"]
    receive_date = extract_result["样品接收时间"]
    detect_type = extract_result["检测类型"]

    # 4. 检查核心信息是否缺失
    if any(v == "未找到对应内容" for v in [customer_name, sample_name, receive_date]):
        print(f"❌ 关键信息缺失，跳过重命名")
        return False

    # 5. 计算过期时间（兼容中英文日期）
    expire_date = calculate_expire_date(receive_date, expire_days)
    if expire_date == "日期解析失败":
        print(f"❌ 过期时间计算失败，跳过重命名")
        return False

    # 6. 拼接新文件名（中英文信息都兼容）
    filename_parts = [customer_name, sample_name, receive_date, f"过期时间({expire_date})"]
    if detect_type:
        filename_parts.append(detect_type)
    new_filename = "_".join(filename_parts) + ".pdf"
    new_filename = filter_invalid_filename_chars(new_filename)

    # 7. 拼接新文件路径
    original_dir = os.path.dirname(original_path)
    new_pdf_path = os.path.join(original_dir, new_filename)

    # 8. 避免覆盖已存在的文件
    if os.path.exists(new_pdf_path):
        print(f"❌ 新文件名已存在，跳过重命名：{new_pdf_path}")
        return False

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
    """批量处理指定目录下的所有PDF文件（兼容中英文模板）"""
    # 统计变量
    total_count = 0
    success_count = 0
    fail_count = 0
    fail_files = []

    # 遍历目录（含子目录）
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

    # 输出汇总
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