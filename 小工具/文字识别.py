import pdfplumber
import re
import os

# ================= 配置 =================
folder_path = r"E:\System\download\厂商ROHS、REACH - 副本"  # 待处理文件夹
output_unmatched_file = os.path.join(folder_path, "未匹配文件.txt")

def clean_filename(text):
    """清理 Windows 文件名非法字符"""
    return re.sub(r'[\\/:*?"<>|]', '', text).strip()

# ================= 定义三套识别方案 =================
schemes = [
    # 中文老版
    {
        "name": "中文老版",
        "fields": {
            "client": ["客户名称"],
            "sample": ["样品名称"],
            "date": ["样品接收时间"]
        },
        "lang": "中"
    },
    # 英文版
    {
        "name": "英文版",
        "fields": {
            "client": ["Client Name"],
            "sample": ["Sample Name"],
            "date": ["Sample Receiving Date", "Sample Received Date", "Receiving Date"]
        },
        "lang": "英"
    },
    # 新增版
    {
        "name": "新增版",
        "fields": {
            "client": ["报告抬头公司名称"],
            "sample": ["样品型号"],
            "date": ["样品接收日期"]
        },
        "lang": "中"  # 可根据需要修改成英文
    }
]

# ================= 批量处理 =================
unmatched_files = []

for root, dirs, files in os.walk(folder_path):
    for f in files:
        if not f.lower().endswith(".pdf"):
            continue
        pdf_path = os.path.join(root, f)
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()
            if not text:
                unmatched_files.append(pdf_path)
                continue

            lines = [l.strip() for l in text.split("\n") if l.strip()]
            matched = False
            result = {}
            used_scheme = None
            lang_suffix = ""

            # 尝试三套方案
            for scheme in schemes:
                temp = {}
                for i, line in enumerate(lines):
                    for field, keys in scheme["fields"].items():
                        if field in temp:
                            continue
                        for key in keys:
                            if key in line:
                                # 1️⃣ 尝试同一行冒号后取值
                                m = re.search(rf"{re.escape(key)}[:：]?\s*(.+)", line)
                                if m and m.group(1).strip():
                                    temp[field] = m.group(1).strip()
                                # 2️⃣ 否则取下一行
                                elif i + 1 < len(lines):
                                    temp[field] = lines[i + 1].strip()
                if len(temp) == 3:
                    result = temp
                    used_scheme = scheme["name"]
                    lang_suffix = scheme["lang"]
                    matched = True
                    break

            if not matched:
                unmatched_files.append(pdf_path)
                continue

            # ================= 重命名 =================
            client = clean_filename(result["client"])
            sample = clean_filename(result["sample"])
            date = clean_filename(result["date"])

            new_name = f"{client}_{sample}_{date}_{lang_suffix}.pdf"
            new_path = os.path.join(root, new_name)

            os.rename(pdf_path, new_path)
            print(f"[重命名成功] {pdf_path} → {new_path}")

        except Exception as e:
            print(f"[处理失败] {pdf_path}，错误：{e}")
            unmatched_files.append(pdf_path)

# ================= 输出未匹配文件 =================
if unmatched_files:
    with open(output_unmatched_file, "w", encoding="utf-8") as f:
        for path in unmatched_files:
            f.write(path + "\n")
    print(f"\n未匹配文件已保存到：{output_unmatched_file}")
else:
    print("\n所有文件均已匹配并重命名成功！")
