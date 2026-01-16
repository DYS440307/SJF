import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse  # 兼容多种日期格式解析

# -------------------------- 全局配置项 --------------------------
# 目标处理目录（所有PDF都在这个目录下，可包含子目录，如需仅单层可修改遍历逻辑）
TARGET_DIR = r'E:\System\download\厂商ROHS、REACH - 副本\1-诚意达\REACH'
# 目标提取项（蓝色框关键词+正则）
target_keys = {
    #SGS中文识别搞定
    "客户名称": r"客户名称[:：]\s*([^\n]+)",
    "样品名称": r"样品名称[:：]\s*([^\n]+)",
    "样品接收时间": r"样品接收时间[:：]\s*([^\n]+)"
}
# 日期格式（提取的时间转datetime用，匹配"2025年05月13日"格式）
date_format = "%Y年%m月%d日"
# 过期时间偏移量（365天）
expire_days = 365
# 要查找的关键字（大小写不敏感）
target_keywords = ["rohs", "reach"]


# -------------------------- 工具函数 --------------------------
def filter_invalid_filename_chars(filename):
    """过滤文件名中的非法字符（Windows系统）"""
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def calculate_expire_date(receive_date_str, date_format, days=365):
    """计算过期时间：接收时间 + 指定天数"""
    try:
        # 解析接收时间为datetime对象
        receive_date = datetime.strptime(receive_date_str, date_format)
        # 计算过期时间
        expire_date = receive_date + timedelta(days=days)
        # 转为和接收时间相同的格式
        return expire_date.strftime(date_format)
    except Exception as e:
        # 兼容其他日期格式（如"2025.5.13"）
        try:
            receive_date = parse(receive_date_str, fuzzy=True)
            expire_date = receive_date + timedelta(days=days)
            return expire_date.strftime(date_format)
        except:
            print(f"⚠️ 日期解析失败：{receive_date_str}，错误：{e}")
            return "日期解析失败"


# -------------------------- 核心提取函数 --------------------------
def pdfplumber_extract_multi_page(pdf_path, target_keys, target_keywords):
    """多页遍历提取原生PDF内容，同时查找指定关键字"""
    extract_result = {key: "未找到对应内容" for key in target_keys}
    extract_result["检测类型"] = ""  # 存储找到的RoHs/REACH关键字
    found_page = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历所有页面（页码从1开始）
            for page_num, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                if not page_text:
                    continue  # 该页无文本，跳过

                # 1. 提取客户名称/样品名称/接收时间（只找还没找到的）
                for key, pattern in target_keys.items():
                    if extract_result[key] == "未找到对应内容":
                        match = re.search(pattern, page_text)
                        if match:
                            extract_result[key] = match.group(1).strip()

                # 2. 查找RoHs/REACH关键字（大小写不敏感，找到第一个即停止）
                if not extract_result["检测类型"]:
                    page_text_lower = page_text.lower()  # 转小写统一匹配
                    for keyword in target_keywords:
                        if keyword in page_text_lower:
                            extract_result["检测类型"] = keyword.upper()  # 转大写拼接
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
    """处理单个PDF文件的重命名，返回处理结果（成功/失败）"""
    print(f"\n========== 开始处理文件：{original_path} ==========")

    # 1. 提取PDF内容（含关键字查找）
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

    # 5. 计算过期时间
    expire_date = calculate_expire_date(receive_date, date_format, expire_days)
    if expire_date == "日期解析失败":
        print(f"❌ 过期时间计算失败，跳过重命名")
        return False

    # 6. 拼接新文件名
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
    """批量处理指定目录下的所有PDF文件"""
    # 统计变量
    total_count = 0  # 总PDF数量
    success_count = 0  # 成功数量
    fail_count = 0  # 失败数量
    fail_files = []  # 失败的文件列表

    # 遍历目录（含子目录，如需仅单层可将os.walk改为os.listdir）
    for root, dirs, files in os.walk(target_dir):
        for file in files:
            # 筛选PDF文件（大小写不敏感）
            if file.lower().endswith(".pdf"):
                total_count += 1
                file_path = os.path.join(root, file)
                # 处理单个文件
                if rename_single_pdf(file_path):
                    success_count += 1
                else:
                    fail_count += 1
                    fail_files.append(file_path)

    # 输出批量处理汇总
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
    # 检查目标目录是否存在
    if not os.path.exists(TARGET_DIR):
        print(f"❌ 目标目录不存在：{TARGET_DIR}")
    else:
        # 执行批量处理
        batch_process_pdfs(TARGET_DIR)