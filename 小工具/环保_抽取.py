import pdfplumber
import re
import os
from datetime import datetime, timedelta

# ================= 配置 =================
folder_path = r"E:\System\download\1-诚意达\REACH"
unmatched_file = os.path.join(folder_path, "未匹配文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")

# ================= 工具函数 =================
def clean_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '', text).strip()

def normalize(text):
    # 去除所有空白字符（空格、tab、全角空格等）并去掉冒号
    return re.sub(r'[\s\u3000]+', '', text).replace(":", "").replace("：", "").lower()

def parse_date(date_str):
    # ★ 修复：支持英文月份
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
        "%b %d, %Y",     # Jun 26, 2024
        "%B %d, %Y"      # July 02, 2024
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

# ================= 成组方案（强绑定） =================
schemes = [
    # 中文方案1
    {
        "lang": "中",
        "fields": {
            "client": ["客户名称"],
            "sample": ["样品名称"],
            "date": ["样品接收时间"]
        }
    },
    # 中文方案2
    {
        "lang": "中",
        "fields": {
            "client": ["报告抬头公司名称"],
            "sample": ["样品型号"],
            "date": ["样品接收日期"]
        }
    },
    # 中文方案3
    {
        "lang": "中",
        "fields": {
            "client": ["报告抬头公司名称"],
            "sample": ["样品名称"],
            "date": ["样品接收日期"]
        }
    },
    # 英文方案
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
                        # ★ 修复：允许 key 后只有空格，没有冒号
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
                text = pdf.pages[0].extract_text()

            if not text:
                unmatched.append(pdf_path)
                continue

            lines = [l.strip() for l in text.split("\n") if l.strip()]

            # ① 特殊整组
            result, lang = try_special_group(lines)

            # ② 普通中文组
            if not result:
                cn_schemes = [s for s in schemes if s["lang"] == "中"]
                result, lang = try_normal_group(cn_schemes, lines)

            # ③ 英文组兜底
            if not result:
                en_schemes = [s for s in schemes if s["lang"] == "英"]
                result, lang = try_normal_group(en_schemes, lines)

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

            # ★ 修复：REACH / SVHC 统一拼 REACH
            if lang == "中":
                keywords = []
                for line in lines:
                    l_lower = line.lower()
                    if 'rohs' in l_lower and 'RoHS' not in keywords:
                        keywords.append('RoHS')
                    if 'reach' in l_lower or 'svhc' in l_lower:
                        if 'REACH' not in keywords:
                            keywords.append('REACH')
                if keywords:
                    new_name += '_' + '_'.join(keywords)

            new_name += ".pdf"

            final_path, is_dup = generate_unique_path(os.path.join(root, new_name))
            os.rename(pdf_path, final_path)

            # ★ 修复：重复文件记录语义
            if is_dup:
                duplicates.append(f"{pdf_path} -> {final_path}")

            print(f"[完成] {final_path}")

        except Exception as e:
            unmatched.append(pdf_path)
            print(f"[异常] {pdf_path} → {e}")

# ================= 输出记录 =================
if unmatched:
    with open(unmatched_file, "w", encoding="utf-8") as f:
        f.write("\n".join(unmatched))

if duplicates:
    with open(duplicate_file, "w", encoding="utf-8") as f:
        f.write("\n".join(duplicates))

print("\n处理完成")
print(f"未匹配：{len(unmatched)}")
print(f"重复命名：{len(duplicates)}")
