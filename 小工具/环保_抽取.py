import pdfplumber
import re
import os
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from tqdm import tqdm
import threading

# ================= 全局配置 =================
folder_path = r"E:\System\download\厂商ROHS、REACH\10-西铭\REACH"
unmatched_file = os.path.join(folder_path, "未匹配文件.txt")
duplicate_file = os.path.join(folder_path, "重复文件.txt")

# 线程安全锁（解决多线程下重复文件名判断+打印混乱问题）
name_lock = threading.Lock()
# 全局已处理文件名集合（多线程共享）
processed_names = set()
# 收集处理结果（仅用于统计）
process_results = []


# ================= 工具函数 =================
def clean_company_name(text):
    """清洗公司名称：仅保留中文部分，剔除英文、地址、数字等冗余内容"""
    if not text:
        return ""
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')
    chinese_parts = chinese_pattern.findall(text)
    if chinese_parts:
        return max(chinese_parts, key=len).strip()
    return text.strip()


def clean_sample_name(text):
    """清洗样品名称：剔除款号、采购订单号、买卖方等冗余内容，保留核心样品名"""
    if not text:
        return ""
    redundant_keywords = [
        "Manufacturer制造商", "Buyer买家", "Style No(s)", "款号",
        "PO No.", "采购订单号", "订单号", "型号", "规格"
    ]
    for keyword in redundant_keywords:
        if keyword in text:
            text = text.split(keyword)[0].strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\u4e00-\u9fff\w\s]', '', text)
    return text.strip()


def clean_filename(text):
    """清理文件名中的非法字符（最终文件名兜底）"""
    if not text:
        return ""
    illegal_chars = r'[\\/:*?"<>|]'
    text = re.sub(illegal_chars, '', text)
    text = text.strip("_ ").strip()
    text = re.sub(r'_+', '_', text)
    return text


def clean_value(val):
    """增强版：清理字段原始值的前缀符号"""
    if not val:
        return ""
    val = val.strip()
    val = re.sub(r'^[\)\s]*[:：]?\s*', '', val)
    val = re.sub(r'(样品名称|Sample Name|Paper body)?\s*[.．-]{2,}\s*', '', val, flags=re.I)
    val = val.strip().strip(".").strip("-").strip()
    return val


def normalize(text):
    """文本归一化：用于字段匹配，不影响最终文件名"""
    if not text:
        return ""
    text = re.sub(r'[\s\u3000\t\n\r]+', '', text)
    text = text.replace(":", "").replace("：", "").replace("．", ".").replace("，", ",")
    text = text.lower()
    text = re.sub(r'([\u4e00-\u9fff])([\u4e00-\u9fff])', r'\1\2', text)
    return text


def parse_date(date_str):
    """解析日期（保持原有逻辑）"""
    if not date_str:
        return None
    date_str = date_str.strip()

    # 中文日期解析
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日?', date_str)
    if m:
        year, month, day = map(int, m.groups())
        return datetime(year, month, day)

    # 兼容只写年月/年份
    m = re.match(r'(\d{4})年(\d{1,2})月', date_str)
    if m:
        year, month = map(int, m.groups())
        return datetime(year, month, 1)
    m = re.match(r'(\d{4})年', date_str)
    if m:
        year = int(m.group(1))
        return datetime(year, 1, 1)

    # 英文/数字解析
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


def is_pdf_valid(pdf_path):
    """校验PDF有效性（保持原有逻辑）"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pdf.pages[0].extract_text()
        return True
    except Exception:
        return False


# ================= 字段匹配规则 =================
schemes = [
    {"lang": "中", "fields": {
        "client": ["Applicant", "申请人公司名称"],
        "sample": ["Sample Description", "样品描述", "Sample(s) received is(are) stated to be", "收到的送测样品为"],
        "date": ["Date of Submission", "样品收取日期"]
    }},
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


# ================= 字段提取逻辑 =================
def extract_field_value(lines, key, lang):
    """提取字段值（保持原有匹配逻辑）"""
    key_n = normalize(key)
    for i, line in enumerate(lines):
        line_n = normalize(line)
        if key_n in line_n and len(key_n) > 1:
            colon_match = re.search(r'[:：]\s*(.+)', line)
            val = ""
            if colon_match and colon_match.group(1).strip():
                val = colon_match.group(1).strip()
            else:
                next_lines = []
                for j in range(1, 4):
                    if i + j < len(lines):
                        next_line = lines[i + j].strip()
                        if next_line and not re.match(r'^[.．-]+$', next_line):
                            next_lines.append(next_line)
                val = " ".join(next_lines)
            val = clean_value(val)
            return val
    return ""


def try_match_scheme(lines, scheme):
    """匹配单个规则（保持原有逻辑）"""
    temp = {}
    for field, keys in scheme["fields"].items():
        for key in keys:
            val = extract_field_value(lines, key, scheme["lang"])
            if val:
                temp[field] = val
                break
    return temp if len(temp) == 3 else None


def try_match_all_schemes(lines):
    """匹配所有规则（保持原有逻辑）"""
    for scheme in schemes:
        result = try_match_scheme(lines, scheme)
        if result:
            return result, scheme["lang"]
    return None, None


# ================= 重复文件处理 =================
def generate_unique_path(base_path):
    """生成唯一路径（保持原有逻辑）"""
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


# ================= 单文件处理函数 =================
def process_single_pdf(pdf_path):
    """处理单个PDF（核心调整HF判定逻辑：仅同时检测到F/Cl/Br/I四种才拼接）"""
    # 加锁避免多线程打印混乱
    with name_lock:
        print(f"\n===== 开始处理文件：{os.path.basename(pdf_path)} =====")

    if 'msds' in pdf_path.lower():
        with name_lock:
            print(f"文件 {os.path.basename(pdf_path)} 包含MSDS，跳过处理")
        return (pdf_path, "", "跳过", "文件包含MSDS，无需处理")

    if not is_pdf_valid(pdf_path):
        with name_lock:
            print(f"文件 {os.path.basename(pdf_path)} 损坏/加密，无法读取")
        return (pdf_path, "", "失败", "PDF文件损坏/加密/无读取权限")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            if not first_page_text:
                with name_lock:
                    print(f"文件 {os.path.basename(pdf_path)} 第一页无文本内容")
                return (pdf_path, "", "失败", "PDF第一页无文本内容")

            first_lines = [l.strip() for l in first_page_text.split("\n") if l.strip()]
            scan_lines = []
            for idx in range(min(2, len(pdf.pages))):
                t = pdf.pages[idx].extract_text()
                if t:
                    scan_lines.extend([l.strip() for l in t.split("\n") if l.strip()])

        # 匹配基础字段
        result, lang = try_match_all_schemes(first_lines)
        if not result:
            with name_lock:
                print(f"文件 {os.path.basename(pdf_path)} 未匹配到公司/样品/日期字段")
            return (pdf_path, "", "失败", "字段匹配失败（无可用规则）")

        # 解析日期
        dt = parse_date(result["date"])
        if not dt:
            with name_lock:
                print(f"文件 {os.path.basename(pdf_path)} 日期解析失败，原始日期：{result['date']}")
            return (pdf_path, "", "失败", f"日期解析失败（原始日期：{result['date']}）")
        expire = dt + timedelta(days=365)

        # 核心：清洗公司名和样品名
        client_raw = result['client']
        sample_raw = result['sample']
        client_clean = clean_company_name(client_raw)
        sample_clean = clean_sample_name(sample_raw)
        client_final = clean_filename(client_clean)
        sample_final = clean_filename(sample_clean)

        with name_lock:
            print(f"清洗后公司名：{client_final}")
            print(f"清洗后样品名：{sample_final}")
            print(f"解析到日期：{dt.strftime('%Y-%m-%d')}，过期时间：{expire.strftime('%Y-%m-%d')}")
            print("----- 开始检测关键词 -----")

        # 关键词识别（新增打印输出）
        keywords = set()
        halogen_hits = set()  # 仅存储F/Cl/Br/I四种卤素元素
        detected_chars = []  # 记录所有检测到的触发字符

        for line in scan_lines:
            l = line.lower()
            # 检测RoHS
            if 'rohs' in l:
                detected_chars.append(f"rohs（行内容：{line[:50]}...）")  # 截取前50字符避免过长
                keywords.add('RoHS')
            # 检测REACH/SVHC
            if 'reach' in l:
                detected_chars.append(f"reach（行内容：{line[:50]}...）")
                keywords.add('REACH')
            if 'svhc' in l:
                detected_chars.append(f"svhc（行内容：{line[:50]}...）")
                keywords.add('REACH')
            # 检测卤素元素（仅F/Cl/Br/I）
            f_match = re.search(r'\bF\b', line, re.I)
            if f_match:
                detected_chars.append(f"F（行内容：{line[:50]}...）")
                halogen_hits.add('F')
            cl_match = re.search(r'\bCl\b', line, re.I)
            if cl_match:
                detected_chars.append(f"Cl（行内容：{line[:50]}...）")
                halogen_hits.add('Cl')
            br_match = re.search(r'\bBr\b', line, re.I)
            if br_match:
                detected_chars.append(f"Br（行内容：{line[:50]}...）")
                halogen_hits.add('Br')
            i_match = re.search(r'\bI\b', line, re.I)
            if i_match:
                detected_chars.append(f"I（行内容：{line[:50]}...）")
                halogen_hits.add('I')
            # 检测中文“卤素”（仅记录，不再作为HF判定条件）
            if '卤素' in line:
                detected_chars.append(f"卤素（行内容：{line[:50]}...）")

        # 打印检测到的字符
        with name_lock:
            if detected_chars:
                print(f"检测到的触发字符：")
                for char in list(set(detected_chars)):  # 去重后打印
                    print(f"  - {char}")
            else:
                print("未检测到任何关键词触发字符")

            # 判定HF并打印逻辑（核心修改：仅同时检测到F/Cl/Br/I四种才拼接）
            print("----- 关键词拼接逻辑 -----")
            if 'RoHS' in keywords:
                print(f"✓ 检测到rohs字符 → 拼接RoHS关键词")
            else:
                print(f"✗ 未检测到rohs字符 → 不拼接RoHS关键词")

            if 'REACH' in keywords:
                print(f"✓ 检测到reach/svhc字符 → 拼接REACH关键词")
            else:
                print(f"✗ 未检测到reach/svhc字符 → 不拼接REACH关键词")

            # HF判定新逻辑：检查是否同时包含F、Cl、Br、I四种元素
            required_halogens = {'F', 'Cl', 'Br', 'I'}
            is_all_halogens_detected = required_halogens.issubset(halogen_hits)
            if is_all_halogens_detected:
                keywords.add('HF')
                print(f"✓ 同时检测到F、Cl、Br、I四种卤素元素（检测到：{halogen_hits}）→ 拼接HF关键词")
            else:
                missing_halogens = required_halogens - halogen_hits
                print(f"✗ 未同时检测到F、Cl、Br、I四种卤素元素 → 不拼接HF关键词")
                print(f"  - 已检测到：{halogen_hits}")
                print(f"  - 缺失：{missing_halogens}")

        # 有序拼接关键词
        keyword_list = []
        if 'RoHS' in keywords:
            keyword_list.append('RoHS')
        if 'REACH' in keywords:
            keyword_list.append('REACH')
        if 'HF' in keywords:
            keyword_list.append('HF')

        # 组装最终文件名
        filename_parts = [
            client_final,
            sample_final,
            dt.strftime('%Y-%m-%d'),
            lang,
            f"过期时间({expire.strftime('%Y-%m-%d')})"
        ]
        if keyword_list:
            filename_parts.append("_".join(keyword_list))

        new_name = "_".join([p for p in filename_parts if p]) + ".pdf"
        new_path = os.path.join(os.path.dirname(pdf_path), new_name)

        # 处理文件名重复
        final_path, is_dup = generate_unique_path(new_path)

        # 重命名文件
        os.rename(pdf_path, final_path)

        with name_lock:
            print(f"最终生成文件名：{os.path.basename(final_path)}")
            print(f"===== 完成处理文件：{os.path.basename(pdf_path)} =====\n")

        return (pdf_path, final_path, "成功" if not is_dup else "重复",
                "文件名重复，自动添加后缀" if is_dup else "处理成功")

    except Exception as e:
        with name_lock:
            print(f"文件 {os.path.basename(pdf_path)} 处理异常：{str(e)}")
        return (pdf_path, "", "失败", f"处理异常：{str(e)}")


# ================= 主流程 =================
def main():
    # 收集PDF文件
    pdf_paths = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_paths.append(os.path.join(root, file))

    if not pdf_paths:
        print("未找到任何PDF文件！")
        return

    # 多线程处理
    max_workers = multiprocessing.cpu_count() * 2
    print(f"\n========== 开始批量处理 ==========")
    print(f"共找到 {len(pdf_paths)} 个PDF文件，启用 {max_workers} 个线程处理")
    print("=================================\n")

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_single_pdf, path): path for path in pdf_paths}
        for future in tqdm(as_completed(futures), total=len(futures), desc="PDF处理进度"):
            process_results.append(future.result())

    # 统计结果
    success = len([r for r in process_results if r[2] == "成功"])
    duplicate = len([r for r in process_results if r[2] == "重复"])
    failed = len([r for r in process_results if r[2] == "失败"])
    skipped = len([r for r in process_results if r[2] == "跳过"])

    # 生成未匹配/重复文件清单
    unmatched = [r[0] for r in process_results if r[2] == "失败" and "字段匹配失败" in r[3]]
    duplicates = [f"{r[0]} -> {r[1]}" for r in process_results if r[2] == "重复"]

    if unmatched:
        with open(unmatched_file, "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched))
    if duplicates:
        with open(duplicate_file, "w", encoding="utf-8") as f:
            f.write("\n".join(duplicates))

    # 打印最终统计
    print(f"\n========== 处理完成统计 ==========")
    print(f"成功重命名：{success} 个")
    print(f"重复文件（自动加后缀）：{duplicate} 个")
    print(f"处理失败：{failed} 个")
    print(f"跳过文件（MSDS）：{skipped} 个")
    if unmatched:
        print(f"未匹配字段文件清单：{unmatched_file}")
    if duplicates:
        print(f"重复文件清单：{duplicate_file}")
    print("=================================")


if __name__ == "__main__":
    main()