import pdfplumber
import re
import os
from datetime import datetime, timedelta

# ================= 配置 =================
folder_path = r"E:\System\download\卤素"
unmatched_file = os.path.join(folder_path, "未匹配文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")


# ================= 工具函数 =================
def clean_filename(text):
    """清理文件名中的非法字符，并去掉首尾空格和下划线"""
    if not text:
        return ""
    text = re.sub(r'[\\/:*?"<>|]', '', text).strip()
    text = text.strip("_ ").strip()
    text = re.sub(r'_+', '_', text)
    return text


def clean_value(val):
    """清理PDF提取字段的前缀，例如 ') : ' 或 ': '"""
    if not val:
        return ""
    val = val.strip()
    val = re.sub(r'^[\)\s]*[:：]?\s*', '', val)  # 去掉前导 ') :' 或 ':'
    return val.strip()


def normalize(text):
    """文本归一化：去除空格、统一符号、转小写"""
    return re.sub(r'[\s\u3000]+', '', text).replace(":", "").replace("：", "").lower()


def parse_date(date_str):
    """增强版日期解析"""
    if not date_str:
        return None
    date_str = re.sub(r'[^0-9a-zA-Z\-/.]', '', date_str.strip())
    date_str = re.sub(r'[/.\s]', '-', date_str)
    date_str = re.sub(r'-+', '-', date_str).strip('-')

    formats = [
        "%d-%b-%Y", "%d-%B-%Y", "%b-%d-%Y", "%B-%d-%Y",
        "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y",
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
        "%y-%m-%d", "%d-%m-%y"
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            if fmt in ["%y-%m-%d", "%d-%m-%y"]:
                dt = dt.replace(year=dt.year + 2000 if dt.year < 100 else dt.year)
            return dt
        except:
            continue

    patterns = [
        r'(\d{4})[-/.\s]*(\d{1,2})[-/.\s]*(\d{1,2})',
        r'(\d{1,2})[-/.\s]*(\d{1,2})[-/.\s]*(\d{4})',
        r'(\d{4})(\d{2})(\d{2})'
    ]

    for pattern in patterns:
        m = re.search(pattern, date_str)
        if m:
            try:
                if pattern in [patterns[0], patterns[2]]:
                    year, month, day = map(int, m.groups())
                else:
                    day, month, year = map(int, m.groups())
                if 1 <= month <= 12 and 1 <= day <= 31:
                    return datetime(year, month, day)
            except:
                continue
    return None


def extract_chinese(text):
    """提取中文连续串"""
    m = re.search(r'[\u4e00-\u9fff]+', text)
    return m.group(0) if m else text


# ================= 成组方案 =================
schemes = [
    {"lang": "中", "fields": {"client": ["客户名称"], "sample": ["样品名称"], "date": ["样品接收时间"]}},
    {"lang": "中", "fields": {"client": ["委托方"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品型号"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "英", "fields": {"client": ["Sample Submitted By"], "sample": ["Sample Name"], "date": ["Sample Receiving Date"]}},
    {"lang": "英", "fields": {"client": ["Client Name"], "sample": ["Sample Name"], "date": ["Sample Receiving Date"]}},
]


# ================= 匹配函数 =================
def try_match_all_schemes(lines):
    for scheme in schemes:
        temp = {}
        i = 0
        while i < len(lines):
            line = lines[i]
            line_n = normalize(line)
            for field, keys in scheme["fields"].items():
                if field in temp:
                    continue
                for key in keys:
                    key_n = normalize(key)
                    if key_n in line_n:
                        m = re.search(rf"{re.escape(key)}\s*[:：]?\s*(.+)", line, re.I)
                        if m and m.group(1).strip():
                            temp[field] = clean_value(m.group(1).strip())
                        elif i + 1 < len(lines):
                            temp[field] = clean_value(lines[i + 1].strip())
                        if scheme["lang"] == "中" and field in temp:
                            temp[field] = extract_chinese(temp[field])
            i += 1
        if len(temp) == 3:
            return temp, scheme["lang"]
    return None, None


# ================= 重复文件生成 =================
processed_names = set()


def generate_unique_path(base_path):
    base, ext = os.path.splitext(base_path)
    name_only = os.path.basename(base_path)
    if name_only not in processed_names:
        processed_names.add(name_only)
        return base_path, False
    i = 1
    while True:
        new_name = f"{base}_重复{i}{ext}"
        name_only_new = os.path.basename(new_name)
        if name_only_new not in processed_names:
            processed_names.add(name_only_new)
            return new_name, True
        i += 1


# ================= 主流程 =================
success_count = 0
duplicates_count = 0
failure_count = 0
failure_reasons = []

unmatched, duplicates = [], []

for root, _, files in os.walk(folder_path):
    for file in files:
        if not file.lower().endswith(".pdf"):
            continue
        pdf_path = os.path.join(root, file)

        try:
            with pdfplumber.open(pdf_path) as pdf:
                first_page_text = pdf.pages[0].extract_text()
                if not first_page_text:
                    unmatched.append(pdf_path)
                    failure_reasons.append("第一页无内容")
                    failure_count += 1
                    continue
                first_lines = [l.strip() for l in first_page_text.split("\n") if l.strip()]

                scan_lines = []
                for idx in range(min(2, len(pdf.pages))):
                    t = pdf.pages[idx].extract_text()
                    if t:
                        scan_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

            result, lang = try_match_all_schemes(first_lines)
            if not result:
                unmatched.append(pdf_path)
                failure_reasons.append("字段匹配失败（无可用方案）")
                failure_count += 1
                continue

            # ===== 调试打印匹配内容 =====
            print("===== 匹配结果 =====")
            for k, v in result.items():
                print(f"{k!r}: {v!r}")

            dt = parse_date(result["date"])
            if not dt:
                unmatched.append(pdf_path)
                failure_reasons.append(f"日期解析失败（原始日期：{result['date']}）")
                failure_count += 1
                continue

            expire = dt + timedelta(days=365)

            # ===== 拼接文件名 =====
            client_clean = clean_filename(result['client']).rstrip("_ ")
            sample_clean = clean_filename(result['sample']).rstrip("_ ")

            parts = [
                client_clean,
                sample_clean,
                dt.strftime('%Y-%m-%d'),
                lang,
                f"过期时间({expire.strftime('%Y-%m-%d')})"
            ]

            # 关键词识别
            keywords = []
            halogen_hits = set()
            for line in scan_lines:
                l = line.lower()
                if 'rohs' in l and 'RoHS' not in keywords:
                    keywords.append('RoHS')
                if 'reach' in l or 'svhc' in l:
                    if 'REACH' not in keywords:
                        keywords.append('REACH')
                if re.search(r'\bF\b', line, re.I): halogen_hits.add('F')
                if re.search(r'\bCl\b', line, re.I): halogen_hits.add('Cl')
                if re.search(r'\bBr\b', line, re.I): halogen_hits.add('Br')
                if re.search(r'\bI\b', line, re.I): halogen_hits.add('I')
            if len(halogen_hits) >= 2: keywords.append('HF')

            if keywords:
                parts.append("_".join(keywords))

            new_name = "_".join(p for p in parts if p) + ".pdf"

            # 处理重复文件
            final_path, is_dup = generate_unique_path(os.path.join(root, new_name))
            os.rename(pdf_path, final_path)

            if is_dup:
                duplicates.append(f"{pdf_path} -> {final_path}")
                duplicates_count += 1
            else:
                success_count += 1

            print(f"[完成] {final_path}")

        except Exception as e:
            unmatched.append(pdf_path)
            failure_reasons.append(f"处理异常：{str(e)}")
            failure_count += 1
            print(f"[异常] {pdf_path} → {e}")

# ================= 输出结果文件 =================
if unmatched:
    with open(unmatched_file, "w", encoding="utf-8") as f:
        for path, reason in zip(unmatched, failure_reasons):
            f.write(f"{path} → {reason}\n")

if duplicates:
    with open(duplicate_file, "w", encoding="utf-8") as f:
        f.write("\n".join(duplicates))

# 打印统计信息
print("\n===== 处理完成 =====")
print(f"成功重命名：{success_count}")
print(f"重复命名（同批次）：{duplicates_count}")
print(f"失败：{failure_count}")
if failure_reasons:
    print("失败原因示例：", failure_reasons[:5])
