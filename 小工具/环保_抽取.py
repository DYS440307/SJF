import pdfplumber
import re
import os
from datetime import datetime, timedelta

# ================= 配置 =================
folder_path = r"E:\System\download\厂商ROHS、REACH\22-金睿得\ROHS"
unmatched_file = os.path.join(folder_path, "未匹配文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")

# ================= 工具函数 =================
def clean_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '', text).strip()

def normalize(text):
    return text.replace(":", "").replace("：", "").replace(" ", "")

def parse_date(date_str):
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%B %d, %Y", "%b %d, %Y",
        "%Y年%m月%d日"
    ):
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
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

# ================= 匹配方案 =================
schemes = [
    # ===== 中文优先组 =====
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

    # ===== 英文兜底组 =====
    {
        "lang": "英",
        "fields": {
            "client": ["Client Name"],
            "sample": ["Sample Name"],
            "date": ["Sample Receiving Date", "Sample Received Date", "Receiving Date"]
        }
    },
    {
        "lang": "英",
        "fields": {
            "client": ["Company Name"],
            "sample": ["Sample Name"],
            "date": ["Sample Received Date"]
        }
    }

]

# ================= 主处理 =================
unmatched = []
duplicates = []

def try_match(scheme_list, lines):
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

            # ===== ① 先跑中文 =====
            cn_schemes = [s for s in schemes if s["lang"] == "中"]
            result, lang = try_match(cn_schemes, lines)

            # ===== ② 再跑英文 =====
            if not result:
                en_schemes = [s for s in schemes if s["lang"] == "英"]
                result, lang = try_match(en_schemes, lines)

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
                f"过期时间({expire.strftime('%Y-%m-%d')}).pdf"
            )

            target_path = os.path.join(root, new_name)
            final_path, is_duplicate = generate_unique_path(target_path)

            os.rename(pdf_path, final_path)

            if is_duplicate:
                duplicates.append(final_path)

            print(f"[完成] {final_path}")

        except Exception as e:
            unmatched.append(pdf_path)
            print(f"[异常] {pdf_path} → {e}")

# ================= 输出结果 =================
if unmatched:
    with open(unmatched_file, "w", encoding="utf-8") as f:
        f.write("\n".join(unmatched))

if duplicates:
    with open(duplicate_file, "w", encoding="utf-8") as f:
        f.write("\n".join(duplicates))

print("\n处理完成")
print(f"未匹配：{len(unmatched)}")
print(f"重复命名：{len(duplicates)}")
