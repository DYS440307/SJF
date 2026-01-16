import pdfplumber
import re
import os
from datetime import datetime, timedelta
from dateutil.parser import parse  # 兼容多种日期格式解析

# -------------------------- 配置项 --------------------------
# 原PDF文件路径
original_pdf_path = r'E:\System\download\厂商ROHS、REACH - 副本\1-诚意达\REACH\东阳市诚意达电子有限公司_盆架_2025年05月13日_2026年05月13日.pdf'
# 目标提取项（蓝色框关键词+正则）
target_keys = {
    #SGS中文解析完成
    "客户名称": r"客户名称[:：]\s*([^\n]+)",
    "样品名称": r"样品名称[:：]\s*([^\n]+)",
    "样品接收时间": r"样品接收时间[:：]\s*([^\n]+)"
}
# 日期格式（提取的时间转datetime用，匹配"2025年05月13日"格式）
date_format = "%Y年%m月%d日"
# 过期时间偏移量（365天）
expire_days = 365


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
            print(f"日期解析失败：{receive_date_str}，错误：{e}")
            return "日期解析失败"


# -------------------------- 核心提取函数 --------------------------
def pdfplumber_extract_multi_page(pdf_path, target_keys):
    """多页遍历提取原生PDF内容"""
    extract_result = {key: "未找到对应内容" for key in target_keys}
    found_page = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 遍历所有页面（页码从1开始）
            for page_num, page in enumerate(pdf.pages, start=1):
                print(f"正在解析第 {page_num} 页...")
                page_text = page.extract_text()
                if not page_text:
                    continue  # 该页无文本，跳过

                # 逐个匹配目标项（只找还没找到的）
                for key, pattern in target_keys.items():
                    if extract_result[key] == "未找到对应内容":
                        match = re.search(pattern, page_text)
                        if match:
                            extract_result[key] = match.group(1).strip()

                # 检查是否所有项都找到，找到则终止遍历
                if all(v != "未找到对应内容" for v in extract_result.values()):
                    found_page = page_num
                    break

        extract_result["找到内容的页码"] = found_page if found_page else "所有页均未找到"
    except Exception as e:
        extract_result = {"error": f"处理失败：{str(e)}"}

    return extract_result


# -------------------------- 文件重命名函数 --------------------------
def rename_pdf_file(original_path, extract_result):
    """根据提取结果重命名PDF文件"""
    # 1. 检查提取结果是否完整
    if "error" in extract_result:
        print(f"提取失败，无法重命名：{extract_result['error']}")
        return False

    customer_name = extract_result["客户名称"]
    sample_name = extract_result["样品名称"]
    receive_date = extract_result["样品接收时间"]

    if any(v == "未找到对应内容" for v in [customer_name, sample_name, receive_date]):
        print("关键信息缺失，无法重命名：")
        print(f"客户名称：{customer_name}，样品名称：{sample_name}，接收时间：{receive_date}")
        return False

    # 2. 计算过期时间
    expire_date = calculate_expire_date(receive_date, date_format, expire_days)
    if expire_date == "日期解析失败":
        return False

    # 3. 拼接新文件名（过滤非法字符）【核心修改处】
    new_filename = f"{customer_name}_{sample_name}_{receive_date}_过期时间({expire_date}).pdf"
    new_filename = filter_invalid_filename_chars(new_filename)

    # 4. 拼接新文件路径（和原文件同目录）
    original_dir = os.path.dirname(original_path)
    new_pdf_path = os.path.join(original_dir, new_filename)

    # 5. 执行重命名（避免覆盖已存在的文件）
    if os.path.exists(new_pdf_path):
        print(f"新文件名已存在：{new_pdf_path}，重命名失败")
        return False

    try:
        os.rename(original_path, new_pdf_path)
        print(f"文件重命名成功！")
        print(f"原路径：{original_path}")
        print(f"新路径：{new_pdf_path}")
        return True
    except Exception as e:
        print(f"重命名失败：{str(e)}")
        return False


# -------------------------- 主执行逻辑 --------------------------
if __name__ == "__main__":
    # 1. 提取PDF内容
    extract_result = pdfplumber_extract_multi_page(original_pdf_path, target_keys)
    print("\n=== 提取结果 ===")
    for key, value in extract_result.items():
        print(f"{key}：{value}")

    # 2. 重命名文件（仅当提取成功时）
    if "error" not in extract_result:
        rename_pdf_file(original_pdf_path, extract_result)