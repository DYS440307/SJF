import openpyxl
import os
import re


def clean_text(text):
    """清洗文本：去除空格、特殊字符并统一为小写，增强匹配度"""
    if not text:
        return ""
    text_str = str(text).strip()
    text_str = re.sub(r'\s+', '', text_str)  # 去除所有空格
    text_str = re.sub(r'[^\w\u4e00-\u9fa5]', '', text_str)  # 保留字母、数字和中文
    return text_str.lower()  # 统一小写，忽略大小写差异


# ========== 文件路径 ==========
file1_path = r"E:\System\desktop\PY\SQE\关系梳理\2_惠州声乐品质履历_IQC检验记录汇总 - 副本.xlsx"
file2_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表_优化 - 副本.xlsx"

# ========== 要操作的月份工作表名 ==========
target_month = "1月"

# ================== 加载文件 ==================
if not os.path.exists(file1_path):
    raise FileNotFoundError(f"文件1不存在：{file1_path}")
if not os.path.exists(file2_path):
    raise FileNotFoundError(f"文件2不存在：{file2_path}")

# 只读模式加载文件1（提高性能）
wb1 = openpyxl.load_workbook(file1_path, read_only=True, data_only=True)
wb2 = openpyxl.load_workbook(file2_path)

# 文件2：定位到指定月份工作表
if target_month not in wb2.sheetnames:
    raise ValueError(f"文件2中未找到工作表：{target_month}")
ws2 = wb2[target_month]

# ================== 读取文件1并统计数据 ==================
ws1 = wb1.active  # 默认激活的工作表
total_count = {}  # 存储总数量统计
ng_count = {}  # 存储NG数量统计
total_processed = 0
valid_count = 0
ng_total = 0

print("开始处理文件1数据...")
# 从第2行开始统计（跳过表头）
for row in ws1.iter_rows(min_row=2, values_only=True):
    total_processed += 1
    # 提取前四列数据（日期、供应商、部品、状态）
    date, supplier, part, status = row[:4]

    # 跳过关键信息为空的行
    if not supplier or not part:
        continue

    # 判断是否为目标月份数据
    is_target_month = False
    if isinstance(date, str):
        if "1月" in date:
            is_target_month = True
    elif hasattr(date, "month"):
        if date.month == 1:
            is_target_month = True

    if not is_target_month:
        continue

    # 清洗文本后生成匹配键
    cleaned_supplier = clean_text(supplier)
    cleaned_part = clean_text(part)
    key = (cleaned_supplier, cleaned_part)

    # 统计总数量
    total_count[key] = total_count.get(key, 0) + 1
    valid_count += 1

    # 统计NG数量（判断状态是否为NG，不区分大小写）
    if status and str(status).strip().lower() == "ng":
        ng_count[key] = ng_count.get(key, 0) + 1
        ng_total += 1

print(f"文件1处理完成：共{total_processed}行，有效数据{valid_count}行，其中NG数据{ng_total}行")
print(f"统计到{len(total_count)}种供应商-部品组合")

# ================== 写入文件2 ==================
print("\n开始写入文件2...")
updated_rows = 0
no_match_rows = 0

# 从第6行开始处理（文件2数据起始行）
for row_num in range(6, ws2.max_row + 1):
    # 获取文件2中的部品和供应商（B列和C列）
    part = ws2.cell(row=row_num, column=2).value
    supplier = ws2.cell(row=row_num, column=3).value

    # 跳过空值行
    if not supplier or not part:
        continue

    # 清洗文本并生成匹配键
    cleaned_supplier = clean_text(supplier)
    cleaned_part = clean_text(part)
    key = (cleaned_supplier, cleaned_part)

    # 写入总数量（第4列）
    total = total_count.get(key, 0)
    ws2.cell(row=row_num, column=4).value = total

    # 写入NG数量（第5列）
    ng = ng_count.get(key, 0)
    ws2.cell(row=row_num, column=5).value = ng

    updated_rows += 1
    # 调试信息（需要时取消注释）
    # print(f"行{row_num}：总数量={total}, NG数量={ng}")

print(f"文件2处理完成：共更新{updated_rows}行数据")

# ================== 保存 ==================
save_path = file2_path.replace(".xlsx", "_更新后.xlsx")
wb2.save(save_path)
wb1.close()  # 关闭只读模式的工作簿
wb2.close()

print(f"\n结果已保存至：{save_path}")
