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
    return text.replace(":", "").replace("：", "").replace(" ", "")

def parse_date(date_str):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(date_str, fmt)
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

def contains_any_chinese_key(lines, schemes):
    for scheme in schemes:
        if scheme["lang"] != "中":
            continue
        for keys in scheme["fields"].values():
            for key in keys:
                for line in lines:
                    if normalize(key) in normalize(line):
                        return True
    return False

# ================= 特殊整组（最高优先） =================
SPECIAL_GROUP = {
    "lang": "中",
    "patterns": {
        "client": r"委托方\s*/\s*Applicant.*?[:：]\s*(.+)",
        "sample": r"样品名称\s*/\s*Sample\s*Name.*?[:：]\s*(.+)",
        "date": r"样品接收日期\s*/\s*Date\s*of\s*Receipt\s*sample.*?[:：]\s*(\d{4}[-/.]\d{2}[-/.]\d{2})"
    }
}

# ================= 成组方案（强绑定） =================
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
        for i, line in enumerate(lines):
            line_n = normalize(line)
            for field, keys in scheme["fields"].items():
                if field in temp:
                    continue
                for key in keys:
                    if normalize(key) in line_n:
                        m = re.search(rf"{re.escape(key)}[:：]?\s*(.+)", line)
                        if m and m.group(1).strip():
                            temp[field] = m.group(1).strip()
                        elif i + 1 < len(lines):
                            temp[field] = lines[i + 1].strip()
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
            has_cn = contains_any_chinese_key(lines, schemes)

            # ① 特殊整组
            result, lang = try_special_group(lines)

            # ② 普通中文组
            if not result:
                cn_schemes = [s for s in schemes if s["lang"] == "中"]
                result, lang = try_normal_group(cn_schemes, lines)

            # ③ 英文兜底（仅当完全无中文）
            if not result and not has_cn:
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

            # ===== 生成基础文件名 =====
            new_name = (
                f"{clean_filename(result['client'])}_"
                f"{clean_filename(result['sample'])}_"
                f"{dt.strftime('%Y-%m-%d')}_{lang}_"
                f"过期时间({expire.strftime('%Y-%m-%d')})"
            )

            # ===== 中文组额外拼接 RoHs / REACH(SVHC) =====
            if lang == "中":
                keywords = []
                for line in lines:
                    l_lower = line.lower()
                    # 识别 RoHs
                    if 'rohs' in l_lower and 'RoHs' not in keywords:
                        keywords.append('RoHs')
                    # 识别 REACH 或 SVHC
                    if ('reach' in l_lower or 'svhc' in l_lower) and 'REACH(SVHC)' not in keywords:
                        keywords.append('REACH(SVHC)')
                if keywords:
                    new_name += '_' + '_'.join(keywords)

            new_name += ".pdf"

            # ===== 处理重复 =====
            final_path, is_dup = generate_unique_path(os.path.join(root, new_name))
            os.rename(pdf_path, final_path)
            if is_dup:
                duplicates.append(final_path)

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
