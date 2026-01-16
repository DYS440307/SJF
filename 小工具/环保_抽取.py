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
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
        "%b %d, %Y", "%B %d, %Y"
    ):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except:
            pass
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

# ================= 特殊整组（最高优先） =================
SPECIAL_GROUP = {
    "lang": "中",
    "patterns": {
        "client": r"委托方\s*/\s*Applicant.*?[:：]\s*(.+)",
        "sample": r"样品名称\s*/\s*Sample\s*Name.*?[:：]\s*(.+)",
        "date": r"样品接收日期\s*/\s*Date\s*of\s*Receipt\s*sample.*?[:：]\s*(.+)"
    }

}

# ================= 成组方案 =================
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
            "client": ["Client Name"],
            "sample": ["Sample Name"],
            "date": ["Sample Receiving Date", "Sample Received Date"]
        }
    }
]

# ================= 匹配函数 =================
def try_special_group(lines):
    temp = {}
    for field, pattern in SPECIAL_GROUP["patterns"].items():
        for line in lines:
            m = re.search(pattern, line)
            if m:
                temp[field] = m.group(1).strip()
                break
    return (temp, SPECIAL_GROUP["lang"]) if len(temp) == 3 else (None, None)

def try_normal_group(scheme_list, lines):
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
                        m = re.search(
                            rf"{re.escape(key)}\s*[:：]?\s*(.+)",
                            line,
                            re.I
                        )
                        if m and m.group(1).strip():
                            temp[field] = m.group(1).strip()
                        elif i + 1 < len(lines):
                            temp[field] = lines[i + 1].strip()
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

                # ===== 第 1 页：字段识别 =====
                first_page_text = pdf.pages[0].extract_text()
                if not first_page_text:
                    unmatched.append(pdf_path)
                    continue

                first_lines = [l.strip() for l in first_page_text.split("\n") if l.strip()]

                # ===== 第 1 + 第 2 页：关键词扫描 =====
                scan_lines = []
                for idx in range(min(2, len(pdf.pages))):
                    t = pdf.pages[idx].extract_text()
                    if t:
                        scan_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

            # ① 特殊整组
            result, lang = try_special_group(first_lines)

            # ② 中文组
            if not result:
                cn_schemes = [s for s in schemes if s["lang"] == "中"]
                result, lang = try_normal_group(cn_schemes, first_lines)

            # ③ 英文组
            if not result:
                en_schemes = [s for s in schemes if s["lang"] == "英"]
                result, lang = try_normal_group(en_schemes, first_lines)

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

            # ===== 通用关键词识别（第 1 + 第 2 页）=====
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
