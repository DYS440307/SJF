import pdfplumber
import re
import os
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from tqdm import tqdm
import threading
from dateutil import parser

# ================= 字段匹配规则 =================
schemes = [
    {"lang": "中", "fields": {"client": ["申请商", "Applicant", "申请人公司名称"], "sample": ["产品名称", "Sample Description", "样品描述","Sample(s) received is(are) stated to be","收到的送测样品为"], "date": ["样 品 接 收 日 期", "Date of Submission", "样品收取日期"]}},
    {"lang": "中", "fields": {"client": ["客户名称"], "sample": ["样品名称"], "date": ["样品接收时间"]}},
    {"lang": "中", "fields": {"client": ["客户名称"], "sample": ["样品名称"], "date": ["收样日期"]}},
    {"lang": "中", "fields": {"client": ["委托方"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品型号"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["委托单位"], "sample": ["材 质"], "date": ["接收日期"]}},
    {"lang": "中", "fields": {"client": ["委托单位"], "sample": ["样品名称"], "date": ["接收日期"]}},
    {"lang": "中", "fields": {"client": ["申请单位"], "sample": ["样品名称"], "date": ["送样日期"]}},
    {"lang": "中", "fields": {"client": ["委托单位"], "sample": ["Sample Name 样品名称"], "date": ["Received Date 接收日期"]}},
    {"lang": "中", "fields": {"client": ["申请商"], "sample": ["产品名称 ProductName"], "date": ["样 品 接 收 日 期"]}},
    {"lang": "英", "fields": {"client": ["Company Name shown on Report", "Company Name"], "sample": ["Sample Name"], "date": ["Sample Received Date"]}},
    {"lang": "英","fields": {"client": ["Sample Submitted By"], "sample": ["Sample Name"], "date": ["Sample Receiving Date"]}},
    {"lang": "英", "fields": {"client": ["Customer"], "sample": ["Materials"], "date": ["SampleReceivedDate"]}},
    {"lang": "英", "fields": {"client": ["ClientName"], "sample": ["SampleName"], "date": ["DateofSampleReceived"]}},
    {"lang": "英", "fields": {"client": ["Applicant"], "sample": ["SampleName"], "date": ["SampleReceivedDate"]}}
]

# ================= 全局配置 =================
folder_path = r"E:\System\download\失效pdf\AAAA"
failed_file = os.path.join(folder_path, "处理失败文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")
DEFAULT_CLIENT = "FengshunShuangxingTechnologyCo ,Ltd"

# ================= 线程锁 =================
name_lock = threading.Lock()
rename_lock = threading.Lock()
process_results = []

# ================= 工具函数 =================
def clean_company_name(text, pdf_filename=""):
    if not text or text.strip() == "":
        return DEFAULT_CLIENT
    english_pattern = re.compile(r'^[A-Za-z0-9\s,.&()-]+$')
    if english_pattern.match(text.strip()):
        return text.strip()
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')
    chinese_parts = chinese_pattern.findall(text)
    if chinese_parts:
        return max(chinese_parts, key=len).strip()
    if "馨固" in pdf_filename:
        return "东莞市馨固电子有限公司"
    cleaned = re.sub(r'[^\u4e00-\u9fffA-Za-z0-9\s,.&()-]', '', text)
    return cleaned.strip() or DEFAULT_CLIENT

def clean_sample_name(text):
    if not text:
        return ""
    text = re.sub(r'^(SampleName|样品名称|产品名称)\s*', '', text, flags=re.I)
    redundant_keywords = ["Manufacturer制造商", "Buyer买家", "Style No(s)", "款号",
                          "PO No.", "采购订单号", "订单号", "型号", "规格",
                          "Color", "Material", "Testing Period"]
    for keyword in redundant_keywords:
        if keyword in text:
            text = text.split(keyword)[0].strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\u4e00-\u9fff\w\s+]', '', text)
    glue_pattern = re.compile(r'胶水|不干胶|米白色胶水', re.I)
    if glue_pattern.search(text):
        return "胶水"
    return text.strip()

def clean_filename(text):
    if not text:
        return ""
    illegal_chars = r'[\\/:*?"<>|]'
    text = re.sub(illegal_chars, '', text)
    text = text.strip("_ ").strip()
    text = re.sub(r'_+', '_', text)
    return text

def clean_value(val):
    if not val:
        return ""
    val = val.strip()
    val = re.sub(r'^[\)\s]*[:：]?\s*', '', val)
    val = re.sub(r'(样品名称|Sample Name|Paper body|Company Name)?\s*[.．-]{2,}\s*', '', val, flags=re.I)
    val = val.strip().strip(".").strip("-").strip()
    return val

def normalize(text):
    if not text:
        return ""
    text = re.sub(r'[\s\u3000\t\n\r]+', '', text)
    text = text.replace(":", "").replace("：", "").replace("．", ".").replace("，", ",")
    return text.lower()

# ================= 日期解析 =================
def parse_date(date_str):
    if not date_str:
        return datetime(2026, 12, 2)
    date_str = date_str.strip()
    date_str_ch = date_str.replace("　", "").replace("年", "-").replace("月", "-").replace("日", "")
    try:
        return parser.parse(date_str_ch, dayfirst=False, fuzzy=True)
    except:
        pass
    date_str_en = re.sub(r'([A-Za-z]{3,9})(\d{1,2})', r'\1 \2', date_str)
    try:
        return parser.parse(date_str_en, fuzzy=True)
    except:
        pass
    try:
        return datetime.strptime(re.sub(r'[^\d]', '', date_str)[:8], '%Y%m%d')
    except:
        pass
    return datetime(2026, 12, 2)

def is_pdf_valid(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pdf.pages[0].extract_text()
        return True
    except:
        return False

def looks_like_date(text):
    if not text:
        return False
    t = text.lower()
    blacklist = ["test requested", "specified by client", "to test", "fluorine", "lead", "cadmium", "mercury"]
    if any(b in t for b in blacklist):
        return False
    return bool(re.search(r"\d{4}|\d{1,2}[-/.]\d{1,2}", t))

# ================= 字段提取 =================
def extract_field_value(lines, key, field_name=None):
    key_n = normalize(key)
    for i, line in enumerate(lines):
        combined_line = line
        if i + 1 < len(lines):
            combined_line += " " + lines[i + 1].strip()
        line_n = normalize(combined_line)
        if re.search(rf'\b{re.escape(key_n)}\b', line_n) and len(key_n) > 1:
            val = ""
            m = re.search(rf"{re.escape(key)}\s*[:：]?\s*(.+)", combined_line, re.I)
            if m and m.group(1).strip():
                val = m.group(1).strip()
            else:
                for j in range(i + 1, min(i + 4, len(lines))):
                    candidate = lines[j].strip()
                    if field_name == "date" and not looks_like_date(candidate):
                        continue
                    val = candidate
                    break
            if field_name == "client":
                val = re.sub(r'(Company Name|Client Name|委托方|委托单位|Applicant)', '', val, flags=re.I)
            elif field_name == "sample":
                val = re.sub(r'(Sample Name|样品名称|样品描述|产品名称)', '', val, flags=re.I)
            return clean_value(val) if val else None
    return None

# ================= 匹配规则 =================
def try_match_scheme(lines, scheme):
    temp = {}
    for field, keys in scheme["fields"].items():
        for key in keys:
            val = extract_field_value(lines, key, field_name=field)
            if val:
                temp[field] = val
                break
    return temp

def try_match_all_schemes(lines):
    """
    尝试匹配所有schemes，返回匹配结果、语言、以及scheme编号
    """
    for idx, scheme in enumerate(schemes, start=1):
        result = try_match_scheme(lines, scheme)
        if result:
            return result, scheme["lang"], idx
    return None, None, None

# ================= 文件重命名 =================
def safe_rename(src, target):
    base, ext = os.path.splitext(target)
    with rename_lock:
        if not os.path.exists(target):
            os.replace(src, target)
            return target, False
        i = 1
        while True:
            new_target = f"{base}_重复{i}{ext}"
            if not os.path.exists(new_target):
                os.replace(src, new_target)
                return new_target, True
            i += 1

# ================= PDF处理 =================
def process_single_pdf(pdf_path):
    pdf_filename = os.path.basename(pdf_path)
    with name_lock:
        print(f"\n===== 开始处理文件：{pdf_filename} =====", flush=True)

    if 'msds' in pdf_path.lower():
        print(f"提取结果 -> client: 未读取, sample: 未读取, date: 未读取, scheme编号: -")
        return (pdf_path, "", "跳过", "文件包含MSDS，无需处理")

    if not is_pdf_valid(pdf_path):
        print(f"提取结果 -> client: 未读取, sample: 未读取, date: 未读取, scheme编号: -")
        return (pdf_path, "", "失败", "PDF文件损坏/加密/无读取权限")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_lines = []
            for idx in range(min(2, len(pdf.pages))):
                t = pdf.pages[idx].extract_text()
                if t:
                    first_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

        result, lang, scheme_index = try_match_all_schemes(first_lines)

        client_val = result.get('client') if result else None
        sample_val = result.get('sample') if result else None
        date_val = result.get('date') if result else None

        if not sample_val:
            for line in first_lines:
                if re.search(r'Sample Name|样品名称|产品名称', line, re.I):
                    sample_val = extract_field_value(first_lines, line, "sample")
                    if sample_val:
                        break
            if not sample_val:
                sample_val = re.sub(r'[^\w+]', '', pdf_filename.split(".")[0])

        if not client_val:
            client_val = DEFAULT_CLIENT
        if not date_val:
            date_val = datetime.now().strftime('%Y-%m-%d')

        # 控制台打印提取结果及scheme编号
        with name_lock:
            print(f"提取结果 -> client: {client_val}, sample: {sample_val}, date: {date_val}, scheme编号: {scheme_index}")

        dt = parse_date(date_val)
        expire = dt + timedelta(days=365)

        client_final = clean_filename(clean_company_name(client_val, pdf_filename))
        sample_final = clean_filename(clean_sample_name(sample_val))

        keywords = set()
        halogen_hits = set()
        for line in first_lines:
            l = line.lower()
            if re.search(r'roh\s*s', l):
                keywords.add('RoHS')
            if 'reach' in l or 'svhc' in l:
                keywords.add('REACH')
            for h in ['F', 'Cl', 'Br', 'I']:
                if re.search(rf'\b{h}\b', line, re.I):
                    halogen_hits.add(h)
        if {'F', 'Cl', 'Br', 'I'}.issubset(halogen_hits):
            keywords.add('HF')
        if not keywords:
            keywords.add('RoHS')

        keyword_list = [k for k in ['RoHS', 'REACH', 'HF'] if k in keywords]

        filename_parts = [
            client_final,
            sample_final,
            dt.strftime('%Y-%m-%d'),
            lang or "英",
            f"过期时间({expire.strftime('%Y-%m-%d')})"
        ]
        if keyword_list:
            filename_parts.append("_".join(keyword_list))

        new_name = "_".join([p for p in filename_parts if p]) + ".pdf"
        new_path = os.path.join(os.path.dirname(pdf_path), new_name)
        final_path, is_dup = safe_rename(pdf_path, new_path)

        return (pdf_path, final_path, "成功" if not is_dup else "重复",
                "处理成功" if not is_dup else "文件名重复，自动添加后缀")

    except Exception as e:
        print(f"提取结果 -> client: 未读取, sample: 未读取, date: 未读取, scheme编号: -")
        return (pdf_path, "", "失败", f"处理异常：{str(e)}")

# ================= 主函数 =================
def main():
    pdf_paths = [os.path.join(root, f) for root, _, files in os.walk(folder_path)
                 for f in files if f.lower().endswith(".pdf")]
    if not pdf_paths:
        print("未找到任何PDF文件！")
        return

    max_workers = multiprocessing.cpu_count() * 2
    print(f"\n========== 开始批量处理 ==========")
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_single_pdf, p): p for p in pdf_paths}
        for f in tqdm(as_completed(futures), total=len(futures), desc="PDF处理进度"):
            process_results.append(f.result())

    success = len([r for r in process_results if r[2] == "成功"])
    duplicate = len([r for r in process_results if r[2] == "重复"])
    failed = len([r for r in process_results if r[2] == "失败"])
    skipped = len([r for r in process_results if r[2] == "跳过"])

    failed_records = [f"{r[0]} -> 失败原因：{r[3]}" for r in process_results if r[2] == "失败"]
    if failed_records:
        with open(failed_file, "w", encoding="utf-8") as f:
            f.write("\n".join(failed_records))

    duplicates = [f"{r[0]} -> 重命名为：{r[1]}" for r in process_results if r[2] == "重复"]
    if duplicates:
        with open(duplicate_file, "w", encoding="utf-8") as f:
            f.write("\n".join(duplicates))

    print(f"\n========== 处理完成统计 ==========")
    print(f"成功重命名：{success} 个")
    print(f"重复文件（自动加后缀）：{duplicate} 个")
    print(f"处理失败：{failed} 个")
    print(f"跳过文件（MSDS）：{skipped} 个")
    if failed_records:
        print(f"处理失败文件清单（含原因）：{failed_file}")
    if duplicates:
        print(f"重复文件清单：{duplicate_file}")
    print("=================================")

if __name__ == "__main__":
    main()
