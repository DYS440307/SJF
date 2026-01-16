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
    return re.sub(r'[\\/:*?"<>|]', '', text).strip()

def normalize(text):
    return re.sub(r'[\s\u3000]+', '', text).replace(":", "").replace("：", "").lower()

def parse_date(date_str):
    """兼容各种日期格式"""
    date_str = date_str.strip()
    date_str = date_str.replace("年", "-").replace("月", "-").replace("日", "")
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d-%b-%Y", "%d-%B-%Y", "%b %d, %Y", "%B %d, %Y",
        "%d-%m-%Y", "%d/%m/%Y"
    ):
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    # 尝试仅用数字提取年月日
    m = re.search(r"(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})", date_str)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

def generate_unique_path(base_path):
    if not os.path.exists(base_path):
        return base_path, False
    base, ext = os.path.splitext(base_path)
    i = 1
    while True:
        new_path = f"{base}_重复{i}{ext}"
        if not os.path.exists(new_path):
            return new_path, True
        i += 1

def extract_chinese(text):
    """提取文本中的中文字符连续串"""
    m = re.search(r'[\u4e00-\u9fff]+', text)
    return m.group(0) if m else text

# ================= 成组方案（保持拆开） =================
schemes = [
    {
        "lang": "中",
        "fields": {
            "client": ["客户名称"],
            "sample": ["样品名称"],
            "date": ["样品接收时间"]
        }
    },
    {
        "lang": "中",
        "fields": {
            "client": ["委托方"],
            "sample": ["样品名称"],
            "date": ["样品接收日期"]
        }
    },
    {
        "lang": "中",
        "fields": {
            "client": ["报告抬头公司名称"],
            "sample": ["样品型号"],
            "date": ["样品接收日期"]
        }
    },
    {
        "lang": "中",
        "fields": {
            "client": ["报告抬头公司名称"],
            "sample": ["样品名称"],
            "date": ["样品接收日期"]
        }
    },
    {
        "lang": "英",
        "fields": {
            "client": ["Sample Submitted By"],
            "sample": ["Sample Name"],
            "date": ["Sample Receiving Date"]
        }
    },
    {
        "lang": "英",
        "fields": {
            "client": ["Client Name"],
            "sample": ["Sample Name"],
            "date": ["Sample Receiving Date"]
        }
    }
]

# ================= 语言识别 =================
def detect_language(lines):
    """判断文本语言类型"""
    has_simplified = has_traditional = has_english = False
    traditional_chars = "電體樣品廠商"  # 常用繁体字
    for line in lines:
        # 中文字符
        if re.search(r'[\u4e00-\u9fff]', line):
            if any(c in line for c in traditional_chars):
                has_traditional = True
            else:
                has_simplified = True
        # 英文
        if re.search(r'[A-Za-z]', line):
            has_english = True

    # 逻辑判断
    if has_simplified and has_english:
        return "中"
    if has_traditional and has_english:
        return "英"
    if has_simplified or has_traditional:
        return "中"
    if has_english:
        return "英"
    return None

# ================= 匹配函数 =================
def try_match(scheme_list, lines):
    """尝试匹配单独组，匹配成功即返回"""
    for scheme in scheme_list:
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
                            temp[field] = m.group(1).strip()
                        elif i + 1 < len(lines):
                            temp[field] = lines[i + 1].strip()
                        # 中文组字段只保留中文
                        if scheme["lang"] == "中":
                            temp[field] = extract_chinese(temp[field])
            i += 1
        if len(temp) == 3:
            return temp, scheme["lang"]
    return None, None

# ================= 主流程 =================
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
                    continue
                first_lines = [l.strip() for l in first_page_text.split("\n") if l.strip()]

                # 扫描第1+2页
                scan_lines = []
                for idx in range(min(2, len(pdf.pages))):
                    t = pdf.pages[idx].extract_text()
                    if t:
                        scan_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

            # ===== 语言判断 =====
            lang_detected = detect_language(first_lines)
            if not lang_detected:
                unmatched.append(pdf_path)
                continue

            # 选择对应语言的独立组
            scheme_list = [s for s in schemes if s["lang"] == lang_detected]

            # ===== 匹配字段 =====
            result, lang = try_match(scheme_list, first_lines)
            if not result:
                unmatched.append(pdf_path)
                continue

            dt = parse_date(result["date"])
            if not dt:
                unmatched.append(pdf_path)
                continue

            expire = dt + timedelta(days=365)

            new_name = (
                f"{clean_filename(result['client'])}_"
                f"{clean_filename(result['sample'])}_"
                f"{dt.strftime('%Y-%m-%d')}_{lang}_"
                f"过期时间({expire.strftime('%Y-%m-%d')})"
            )

            # ===== 通用关键词识别 =====
            keywords = []
            halogen_hits = set()

            for line in scan_lines:
                l = line.lower()
                if 'rohs' in l and 'RoHS' not in keywords:
                    keywords.append('RoHS')
                if 'reach' in l or 'svhc' in l:
                    if 'REACH' not in keywords:
                        keywords.append('REACH')
                # HF 元素识别
                if re.search(r'\bF\b', line, re.I):
                    halogen_hits.add('F')
                if re.search(r'\bCl\b', line, re.I):
                    halogen_hits.add('Cl')
                if re.search(r'\bBr\b', line, re.I):
                    halogen_hits.add('Br')
                if re.search(r'\bI\b', line, re.I):
                    halogen_hits.add('I')

            if len(halogen_hits) >= 2:
                keywords.append('HF')

            if keywords:
                new_name += "_" + "_".join(keywords)

            new_name += ".pdf"

            final_path, is_dup = generate_unique_path(os.path.join(root, new_name))
            os.rename(pdf_path, final_path)
            if is_dup:
                duplicates.append(f"{pdf_path} -> {final_path}")

            print(f"[完成] {final_path}")

        except Exception as e:
            unmatched.append(pdf_path)
            print(f"[异常] {pdf_path} → {e}")

# ================= 输出 =================
if unmatched:
    with open(unmatched_file, "w", encoding="utf-8") as f:
        f.write("\n".join(unmatched))

if duplicates:
    with open(duplicate_file, "w", encoding="utf-8") as f:
        f.write("\n".join(duplicates))

print("\n处理完成")
print(f"未匹配：{len(unmatched)}")
print(f"重复命名：{len(duplicates)}")
