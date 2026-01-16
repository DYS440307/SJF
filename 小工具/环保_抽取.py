import pdfplumber
import re
import os
import shutil
import csv
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from tqdm import tqdm
import threading

# ================= 全局配置 =================
folder_path = r"E:\System\download\厂商ROHS、REACH"
unmatched_file = os.path.join(folder_path, "未匹配文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")
processed_log = os.path.join(folder_path, "已处理文件记录.txt")
report_file = os.path.join(folder_path, "处理报告.csv")
backup_folder = os.path.join(folder_path, "原文件备份")

# 线程安全锁（解决多线程下重复文件名判断的并发问题）
name_lock = threading.Lock()
# 全局已处理文件名集合（多线程共享）
processed_names = set()
# 收集处理结果（用于生成报告）
process_results = []


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
    """增强版：清理PDF提取字段的前缀，过滤冗余字符"""
    if not val:
        return ""
    val = val.strip()
    # 第一步：移除前缀的符号/空格
    val = re.sub(r'^[\)\s]*[:：]?\s*', '', val)
    # 第二步：移除键名相关的冗余字符（如Sample Name、样品名称、连续点/横线）
    val = re.sub(r'(样品名称|Sample Name|Paper body)?\s*[.．-]{2,}\s*', '', val, flags=re.I)
    # 第三步：移除首尾无关字符
    val = val.strip().strip(".").strip("-").strip()
    return val


def normalize(text):
    """增强版文本归一化：应对PDF文字提取偏差（模糊匹配核心）"""
    if not text:
        return ""
    # 1. 移除所有空格、全角空格、制表符、换行符
    text = re.sub(r'[\s\u3000\t\n\r]+', '', text)
    # 2. 统一中英文符号
    text = text.replace(":", "").replace("：", "").replace("．", ".").replace("，", ",")
    # 3. 转小写（消除大小写影响）
    text = text.lower()
    # 4. 处理常见的文字拆分偏差（比如“样品 名称”→“样品名称”、“委托 方”→“委托方”）
    text = re.sub(r'([\u4e00-\u9fff])([\u4e00-\u9fff])', r'\1\2', text)
    return text


def parse_date(date_str):
    """增强版日期解析，支持中文、英文、数字日期"""
    if not date_str:
        return None
    date_str = date_str.strip()

    # ===== 中文日期解析 =====
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日?', date_str)
    if m:
        year, month, day = map(int, m.groups())
        return datetime(year, month, day)

    # 兼容只写年月
    m = re.match(r'(\d{4})年(\d{1,2})月', date_str)
    if m:
        year, month = map(int, m.groups())
        return datetime(year, month, 1)

    # 兼容只写年份
    m = re.match(r'(\d{4})年', date_str)
    if m:
        year = int(m.group(1))
        return datetime(year, 1, 1)

    # ===== 英文/数字解析 =====
    date_str = re.sub(r'[^0-9a-zA-Z\-/.]', '-', date_str)
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

    # 纯数字兜底
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


def is_pdf_valid(pdf_path):
    """校验PDF是否可读、未加密、未损坏（健壮性增强）"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pdf.pages[0].extract_text()
        return True
    except Exception:
        return False


def backup_file(src_path):
    """备份原文件到备份文件夹（可选，防止误操作）"""
    os.makedirs(backup_folder, exist_ok=True)
    rel_path = os.path.relpath(src_path, folder_path)
    dst_path = os.path.join(backup_folder, rel_path)
    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    shutil.copy2(src_path, dst_path)


# ================= 字段匹配规则 =================
schemes = [
    {"lang": "中", "fields": {"client": ["客户名称"], "sample": ["样品名称"], "date": ["样品接收时间"]}},
    {"lang": "中", "fields": {"client": ["客户名称"], "sample": ["样品名称"], "date": ["收样日期"]}},
    {"lang": "中", "fields": {"client": ["委托方"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品型号"], "date": ["样品接收日期"]}},
    {"lang": "中", "fields": {"client": ["报告抬头公司名称"], "sample": ["样品名称"], "date": ["样品接收日期"]}},
    {"lang": "英",
     "fields": {"client": ["Sample Submitted By"], "sample": ["Sample Name"], "date": ["Sample Receiving Date"]}},
    {"lang": "英", "fields": {"client": ["Client Name"], "sample": ["Sample Name"], "date": ["Sample Receiving Date"]}},
    {"lang": "中", "fields": {"client": ["委托单位"], "sample": ["材 质"], "date": ["接收日期"]}}
]


# ================= 字段提取逻辑（模糊匹配） =================
def extract_field_value(lines, key, lang):
    """单独提取单个字段的值（模糊匹配核心）"""
    key_n = normalize(key)
    for i, line in enumerate(lines):
        line_n = normalize(line)
        # 模糊匹配：行内包含键名即可，且键名长度>1（避免空/短键名误匹配）
        if key_n in line_n and len(key_n) > 1:
            # 优先匹配冒号/中文冒号后的内容
            colon_match = re.search(r'[:：]\s*(.+)', line)
            val = ""
            if colon_match and colon_match.group(1).strip():
                val = colon_match.group(1).strip()
            else:
                # 读取后续3行，跳过空行/仅含符号的行
                next_lines = []
                for j in range(1, 4):
                    if i + j < len(lines):
                        next_line = lines[i + j].strip()
                        if next_line and not re.match(r'^[.．-]+$', next_line):
                            next_lines.append(next_line)
                val = " ".join(next_lines)

            val = clean_value(val)
            # 中文字段仅对client提取中文
            if lang == "中" and "client" in key:
                val = extract_chinese(val)
            return val
    return ""


def try_match_scheme(lines, scheme):
    """匹配单个规则方案"""
    temp = {}
    for field, keys in scheme["fields"].items():
        for key in keys:
            val = extract_field_value(lines, key, scheme["lang"])
            if val:
                temp[field] = val
                break
    return temp if len(temp) == 3 else None


def try_match_all_schemes(lines):
    """匹配所有规则方案"""
    for scheme in schemes:
        result = try_match_scheme(lines, scheme)
        if result:
            return result, scheme["lang"]
    return None, None


# ================= 重复文件处理（线程安全） =================
def generate_unique_path(base_path):
    """线程安全的唯一路径生成（加锁防止并发问题）"""
    with name_lock:
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


# ================= 单文件处理函数（适配多线程） =================
def process_single_pdf(pdf_path):
    """
    处理单个PDF文件（多线程核心函数）
    返回值：(原路径, 新路径, 处理状态, 备注)
    状态：成功/重复/失败
    """
    # 跳过MSDS文件
    if 'msds' in pdf_path.lower():
        return (pdf_path, "", "跳过", "文件包含MSDS，无需处理")

    # 校验PDF有效性
    if not is_pdf_valid(pdf_path):
        return (pdf_path, "", "失败", "PDF文件损坏/加密/无读取权限")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            if not first_page_text:
                return (pdf_path, "", "失败", "PDF第一页无文本内容")

            first_lines = [l.strip() for l in first_page_text.split("\n") if l.strip()]
            scan_lines = []
            for idx in range(min(2, len(pdf.pages))):
                t = pdf.pages[idx].extract_text()
                if t:
                    scan_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

        # 匹配字段（模糊匹配）
        result, lang = try_match_all_schemes(first_lines)
        if not result:
            return (pdf_path, "", "失败", "字段匹配失败（无可用规则）")

        # 解析日期
        dt = parse_date(result["date"])
        if not dt:
            return (pdf_path, "", "失败", f"日期解析失败（原始日期：{result['date']}）")

        # 计算过期时间
        expire = dt + timedelta(days=365)

        # 拼接新文件名
        client_clean = clean_filename(result['client']).rstrip("_ ")
        sample_clean = clean_filename(result['sample']).rstrip("_ ")
        parts = [
            client_clean,
            sample_clean,
            dt.strftime('%Y-%m-%d'),
            lang,
            f"过期时间({expire.strftime('%Y-%m-%d')})"
        ]

        # 关键词识别（RoHS/REACH/HF）
        keywords = []
        halogen_hits = set()
        for line in scan_lines:
            l = line.lower()
            if 'rohs' in l and 'RoHS' not in keywords:
                keywords.append('RoHS')
            if 'reach' in l or 'svhc' in l and 'REACH' not in keywords:
                keywords.append('REACH')
            if re.search(r'\bF\b', line, re.I): halogen_hits.add('F')
            if re.search(r'\bCl\b', line, re.I): halogen_hits.add('Cl')
            if re.search(r'\bBr\b', line, re.I): halogen_hits.add('Br')
            if re.search(r'\bI\b', line, re.I): halogen_hits.add('I')
        if len(halogen_hits) >= 2:
            keywords.append('HF')

        if keywords:
            parts.append("_".join(keywords))

        new_name = "_".join(p for p in parts if p) + ".pdf"
        new_path = os.path.join(os.path.dirname(pdf_path), new_name)

        # 生成唯一路径（处理重复）
        final_path, is_dup = generate_unique_path(new_path)

        # 备份原文件（可选，注释掉则关闭备份）
        backup_file(pdf_path)

        # 重命名文件
        os.rename(pdf_path, final_path)

        # 返回结果
        status = "重复" if is_dup else "成功"
        remark = "文件名重复，自动添加后缀" if is_dup else "处理成功"
        return (pdf_path, final_path, status, remark)

    except Exception as e:
        return (pdf_path, "", "失败", f"处理异常：{str(e)}")


# ================= 主流程（多线程+进度条） =================
def main():
    # 1. 初始化文件夹
    os.makedirs(backup_folder, exist_ok=True)

    # 2. 收集所有需要处理的PDF文件
    pdf_paths = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_paths.append(os.path.join(root, file))

    if not pdf_paths:
        print("未找到任何PDF文件！")
        return

    # 3. 配置多线程（按CPU核心数设置，避免线程过多）
    max_workers = multiprocessing.cpu_count() * 2
    print(f"开始处理，共找到 {len(pdf_paths)} 个PDF文件，启用 {max_workers} 个线程...")

    # 4. 多线程处理 + 进度条
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        futures = {executor.submit(process_single_pdf, path): path for path in pdf_paths}
        # 遍历结果（带进度条）
        for future in tqdm(as_completed(futures), total=len(futures), desc="PDF处理进度"):
            result = future.result()
            process_results.append(result)

    # 5. 统计结果
    success = len([r for r in process_results if r[2] == "成功"])
    duplicate = len([r for r in process_results if r[2] == "重复"])
    failed = len([r for r in process_results if r[2] == "失败"])
    skipped = len([r for r in process_results if r[2] == "跳过"])

    # 6. 生成处理报告（CSV）
    with open(report_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["原文件路径", "新文件路径", "处理状态", "备注"])
        writer.writerows(process_results)

    # 7. 生成未匹配/重复文件清单
    unmatched = [r[0] for r in process_results if r[2] == "失败" and "字段匹配失败" in r[3]]
    duplicates = [f"{r[0]} -> {r[1]}" for r in process_results if r[2] == "重复"]

    if unmatched:
        with open(unmatched_file, "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched))

    if duplicates:
        with open(duplicate_file, "w", encoding="utf-8") as f:
            f.write("\n".join(duplicates))

    # 8. 打印统计信息
    print("\n===== 处理完成 =====")
    print(f"成功重命名：{success}")
    print(f"重复文件（自动加后缀）：{duplicate}")
    print(f"处理失败：{failed}")
    print(f"跳过文件（MSDS）：{skipped}")
    print(f"处理报告已导出：{report_file}")
    if unmatched:
        print(f"未匹配文件清单：{unmatched_file}")
    if duplicates:
        print(f"重复文件清单：{duplicate_file}")


if __name__ == "__main__":
    main()