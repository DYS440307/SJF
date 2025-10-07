import openpyxl
import os

# ==================== 配置区 ====================
source_path = r"E:\System\download\2023年.xlsx"   # 原始文件路径
save_path = r"E:\System\download\合并.xlsx"        # 合并后文件路径
target_sheet_name = "合并"                         # 新建工作表名称
# ===============================================

if not os.path.exists(source_path):
    print(f"❌ 找不到文件：{source_path}")
else:
    # 打开源文件
    wb = openpyxl.load_workbook(source_path)
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = target_sheet_name

    first = True  # 控制是否写入表头
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"📄 正在读取工作表：{sheet_name}")

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue  # 跳过空表

        # 第一张表保留表头，其余的去掉表头
        if first:
            for row in rows:
                new_ws.append(row)
            first = False
        else:
            for row in rows[1:]:
                new_ws.append(row)

    # 保存文件
    new_wb.save(save_path)
    print(f"✅ 所有工作表已首尾合并，保存到：{save_path}")
