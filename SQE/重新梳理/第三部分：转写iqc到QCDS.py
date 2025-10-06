import openpyxl
import os

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

wb1 = openpyxl.load_workbook(file1_path)
wb2 = openpyxl.load_workbook(file2_path)

# 文件2：定位到指定月份工作表
if target_month not in wb2.sheetnames:
    raise ValueError(f"文件2中未找到工作表：{target_month}")

ws2 = wb2[target_month]

# ================== 读取文件1并统计组合出现次数 ==================
ws1 = wb1.active  # 默认只有一个sheet或当前激活的

count_dict = {}

# 假设第1行为表头，从第2行开始统计
for row in ws1.iter_rows(min_row=2, values_only=True):
    date, supplier, part, status = row[:4]

    # 跳过空行
    if not supplier or not part:
        continue

    # 判断日期是否为“1月”数据
    # 如果日期是datetime类型，则用 .month 判断；否则用字符串包含判断
    if isinstance(date, str):
        if "1月" not in date:
            continue
    elif hasattr(date, "month"):
        if date.month != 1:
            continue

    key = (supplier.strip(), part.strip())
    count_dict[key] = count_dict.get(key, 0) + 1

# ================== 写入文件2 ==================
# 文件2中从第6行开始，第2列为部品名称，第3列为供应商名称，第4列为总数
for row in range(6, ws2.max_row + 1):
    part = ws2.cell(row=row, column=2).value
    supplier = ws2.cell(row=row, column=3).value

    if not supplier or not part:
        continue

    key = (supplier.strip(), part.strip())

    if key in count_dict:
        ws2.cell(row=row, column=4).value = count_dict[key]
    else:
        ws2.cell(row=row, column=4).value = 0  # 没匹配到的填0

# ================== 保存 ==================
save_path = file2_path.replace(".xlsx", "_更新后.xlsx")
wb2.save(save_path)
print(f"统计完成，结果已保存至：{save_path}")
